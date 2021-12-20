Option Strict On
Option Explicit On

Public Class EstudoIntervalosAcordes

    Private lista As New List(Of Keys)
    Dim FamiliaAcorde As String
    Dim Fonte As New Font("Arial", 12, FontStyle.Italic, GraphicsUnit.Pixel)
    Dim Fonte2 As New Font("Arial", 8, FontStyle.Regular, GraphicsUnit.Pixel)
    Dim Fonte3 As New Font("Times New Roman", 9, FontStyle.Regular, GraphicsUnit.Pixel)
    Dim Fonte4 As New Font("Times New Roman", 20, FontStyle.Bold, GraphicsUnit.Pixel)
    Dim Fonte5 As New Font("Arial", 11, FontStyle.Italic, GraphicsUnit.Pixel)
    Dim Fonte6 As New Font("Times New Roman", 11, FontStyle.Regular, GraphicsUnit.Pixel)
    Dim Fonte7 As New Font("Arial", 9, FontStyle.Regular, GraphicsUnit.Pixel)
    Dim CorFonte As SolidBrush = New SolidBrush(Color.FromArgb(0, 0, 0))
    Dim CorFonte2 As SolidBrush = New SolidBrush(Color.FromArgb(0, 0, 0))
    Dim CorFonte3 As SolidBrush = New SolidBrush(Color.FromArgb(34, 177, 76))
    Dim CorFonte4 As SolidBrush = New SolidBrush(Color.FromArgb(255, 255, 255))
    Dim CorFonte5 As SolidBrush = New SolidBrush(Color.FromArgb(150, 150, 150))
    Dim Cor As SolidBrush
    Dim CorBolinha As SolidBrush = New SolidBrush(Color.FromArgb(0, 0, 0))
    Dim CorBolinhaIntervalo As SolidBrush
    Dim CorBolinhaIntervaloT As SolidBrush = New SolidBrush(Color.FromArgb(0, 0, 0))
    Dim CorBolinhaIntervalo2M As SolidBrush = New SolidBrush(Color.FromArgb(255, 255, 0))
    Dim CorBolinhaIntervalo3M As SolidBrush = New SolidBrush(Color.FromArgb(255, 0, 0))
    Dim CorBolinhaIntervalo4J As SolidBrush = New SolidBrush(Color.FromArgb(0, 204, 255))
    Dim CorBolinhaIntervalo5J As SolidBrush = New SolidBrush(Color.FromArgb(255, 204, 0))
    Dim CorBolinhaIntervalo6M As SolidBrush = New SolidBrush(Color.FromArgb(0, 128, 0))
    Dim CorBolinhaIntervalo7M As SolidBrush = New SolidBrush(Color.FromArgb(255, 153, 204))
    Dim CorPestana As Pen = New Pen(Color.FromArgb(0, 0, 0), 3)
    Dim Seta As Pen = New Pen(Color.FromArgb(0, 0, 0), 5)
    Dim Rect1 As Pen = New Pen(Color.FromArgb(200, 255, 0, 0), 2)
    Dim Rect2 As Pen = New Pen(Color.FromArgb(0, 0, 0), 4)
    Dim PosiçãoNotaFundamental As Pen = New Pen(Color.FromArgb(255, 0, 0), 1)
    Dim RespostaIntervalo As Pen = New Pen(Color.FromArgb(255, 0, 0), 2)
    Dim LinhaBracinhoViolão1A As Pen = New Pen(Color.FromArgb(200, 200, 200), 3)
    Dim LinhaBracinhoViolão1B As Pen = New Pen(Color.FromArgb(153, 217, 234), 3)
    Dim LinhaBracinhoViolão2 As Pen = New Pen(Color.FromArgb(200, 200, 200), 1)
    Dim LinhaDivisória As Pen = New Pen(Color.FromArgb(0, 0, 0), 2)
    Dim Linha As Pen
    Dim LinhaCordas As Pen = New Pen(Color.FromArgb(0, 0, 0), 1)
    Dim LinhaIntervalo As Pen = New Pen(Color.FromArgb(100, 100, 100), 1)
    Dim LinhaIntervalo2 As Pen = New Pen(Color.FromArgb(255, 0, 0), 2)

    Dim NúmeroAcorde, AjusteLeft(6), AjusteTopo, AjustePosição(660), Index, FatorEscala, d, f, Pestana, LinhaFinalDaSeta(660), NumeraçãoTrastes, NumeraçãoFamiliaAcorde, AjusteLeftIntervalos, ValorEsquerda, ValorTopo, TônicaX, TônicaY, PróximaNotaX, PróximaNotaY, DistânciaIntervalo As Integer

    Dim ArrayAcordePadrão(660, 8, 6), ImagemAcorde(660), IntervalosAcorde(27, 6), StringSubstituição(1), IntervalosDoAcordeGerado(6), LocalizaçãoIntervalosNoBraçoDoViolão(20, 6), Baixo, NomeIntervalo As String

    Dim ImagemCopiada(2) As Bitmap

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

    '   Dim soma, contador As Integer

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
        gr.DrawString(txt, Fonte6, CorFonte, x, y, sf)
        sf.Dispose()
    End Sub

    Private Sub EstudoIntervalosAcordes_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Width = My.Settings.NovoValorLarguraTelaEstudoIntervalos
        GerarAcordes()
        Desenhar()
    End Sub

    Private Sub Desenhar()

        Try

            FatorEscala = 1

            ' Make a StringFormat object that centers.
            Dim sf As New StringFormat
            sf.LineAlignment = StringAlignment.Center
            sf.Alignment = StringAlignment.Center

            Dim gr As Graphics
            Dim bmp As New Bitmap(500 * FatorEscala, Me.Height * FatorEscala)
            gr = Graphics.FromImage(bmp)

            'gr.ScaleTransform(FatorEscala, FatorEscala)

            gr.FillRectangle(Brushes.White, 0, 0, 500, Me.Height)

            Dim gradiente As LinearGradientBrush = New LinearGradientBrush(New Rectangle(0, 160, 500, Me.Height - 160), Color.White, Color.LightBlue, LinearGradientMode.Vertical)
            gr.FillRectangle(gradiente, 0, 160, 500, Me.Height - 160)


            gr.TextRenderingHint = TextRenderingHint.ClearTypeGridFit

            Seta.StartCap = LineCap.ArrowAnchor
            PosiçãoNotaFundamental.DashStyle = DashStyle.Dash


            LinhaCordas.CustomEndCap = New AdjustableArrowCap(4, 4, True)
            LinhaCordas.DashStyle = DashStyle.Dot

            LinhaIntervalo.CustomEndCap = New AdjustableArrowCap(4, 4, True)
            LinhaIntervalo.DashStyle = DashStyle.Dot

            LinhaIntervalo2.CustomEndCap = New AdjustableArrowCap(4, 4, True)
            LinhaIntervalo2.DashStyle = DashStyle.Dot

            gr.SmoothingMode = SmoothingMode.AntiAlias



            AjusteTopo = -155

            Linha = LinhaBracinhoViolão1A

            AjusteLeft(0) = 229 : AjusteLeft(1) = 234

            gr.SmoothingMode = SmoothingMode.None
            'desenha os bracinhos do violão
            'linhas horizontais
            For i = 0 To 19
                gr.DrawLine(LinhaBracinhoViolão2, 53 + AjusteLeft(0), (195 + (i * 15)) + AjusteTopo, 128 + AjusteLeft(0), (195 + (i * 15)) + AjusteTopo)

                AjusteLeft(5) = 0
                If 1 + i < 10 Then
                    AjusteLeft(5) = 4
                End If
                gr.DrawString(CStr(1 + i) & ".......", Fonte2, CorFonte5, 25 + AjusteLeft(1) + AjusteLeft(5) + AjusteLeft(6), 183 + AjusteTopo + (i * 15))
            Next

            'linhas verticais
            For i = 0 To 5
                gr.DrawLine(LinhaBracinhoViolão2, (53 + (i * 15)) + AjusteLeft(0), 179 + AjusteTopo, (53 + (i * 15)) + AjusteLeft(0), 480 + AjusteTopo)
            Next
            gr.DrawLine(Linha, 53 + AjusteLeft(0), 180 + AjusteTopo, 129 + AjusteLeft(0), 180 + AjusteTopo) 'esta linha horizontal precisa ser desenhada por último
            gr.DrawLine(Pens.Black, 53 + AjusteLeft(0), (195 + (11 * 15)) + AjusteTopo, 128 + AjusteLeft(0), (195 + (11 * 15)) + AjusteTopo)
            gr.SmoothingMode = SmoothingMode.AntiAlias

            CenterTextAt(gr, "Localização dos Intervalos", 85 + AjusteLeft(0), 174 + AjusteTopo)
            CenterTextAt(gr, "Legenda:", 180 + AjusteLeft(0), 174 + AjusteTopo)

            gr.FillEllipse(CorBolinhaIntervaloT, 159 + AjusteLeft(0), 181 + AjusteTopo, 13, 13)
            gr.DrawEllipse(Pens.Black, 159 + AjusteLeft(0), 181 + AjusteTopo, 13, 13)
            gr.DrawString("Tônica", Fonte6, CorFonte2, 173 + AjusteLeft(0), 181 + AjusteTopo)
            gr.FillEllipse(CorBolinhaIntervalo2M, 159 + AjusteLeft(0), 196 + AjusteTopo, 13, 13)
            gr.DrawEllipse(Pens.Black, 159 + AjusteLeft(0), 196 + AjusteTopo, 13, 13)
            gr.DrawString("Segunda Maior", Fonte6, CorFonte2, 173 + AjusteLeft(0), 196 + AjusteTopo)
            gr.FillEllipse(CorBolinhaIntervalo3M, 159 + AjusteLeft(0), 211 + AjusteTopo, 13, 13)
            gr.DrawEllipse(Pens.Black, 159 + AjusteLeft(0), 211 + AjusteTopo, 13, 13)
            gr.DrawString("Terça Maior", Fonte6, CorFonte2, 173 + AjusteLeft(0), 211 + AjusteTopo)
            gr.FillEllipse(CorBolinhaIntervalo4J, 159 + AjusteLeft(0), 226 + AjusteTopo, 13, 13)
            gr.DrawEllipse(Pens.Black, 159 + AjusteLeft(0), 226 + AjusteTopo, 13, 13)
            gr.DrawString("Quarta Justa", Fonte6, CorFonte2, 173 + AjusteLeft(0), 226 + AjusteTopo)
            gr.FillEllipse(CorBolinhaIntervalo5J, 159 + AjusteLeft(0), 241 + AjusteTopo, 13, 13)
            gr.DrawEllipse(Pens.Black, 159 + AjusteLeft(0), 241 + AjusteTopo, 13, 13)
            gr.DrawString("Quinta Justa", Fonte6, CorFonte2, 173 + AjusteLeft(0), 241 + AjusteTopo)
            gr.FillEllipse(CorBolinhaIntervalo6M, 159 + AjusteLeft(0), 256 + AjusteTopo, 13, 13)
            gr.DrawEllipse(Pens.Black, 159 + AjusteLeft(0), 256 + AjusteTopo, 13, 13)
            gr.DrawString("Sexta Maior", Fonte6, CorFonte2, 173 + AjusteLeft(0), 256 + AjusteTopo)
            gr.FillEllipse(CorBolinhaIntervalo7M, 159 + AjusteLeft(0), 271 + AjusteTopo, 13, 13)
            gr.DrawEllipse(Pens.Black, 159 + AjusteLeft(0), 271 + AjusteTopo, 13, 13)
            gr.DrawString("Sétima Maior", Fonte6, CorFonte2, 173 + AjusteLeft(0), 271 + AjusteTopo)

            'Desenha a localização dos intervalos no braço do violão
            For b = 0 To 6  'percorre as 6 colunas de cada array
                For c = 0 To 20 'percorre as 12 linhas do array
                    AjusteLeftIntervalos = 0
                    If LocalizaçãoIntervalosNoBraçoDoViolão(c, b) <> "" Then
                        If LocalizaçãoIntervalosNoBraçoDoViolão(c, b) = "T" Then
                            CorBolinhaIntervalo = CorBolinhaIntervaloT
                            AjusteLeftIntervalos = -1
                        ElseIf LocalizaçãoIntervalosNoBraçoDoViolão(c, b) = "2" Then
                            CorBolinhaIntervalo = CorBolinhaIntervalo2M
                        ElseIf LocalizaçãoIntervalosNoBraçoDoViolão(c, b) = "3" Then
                            CorBolinhaIntervalo = CorBolinhaIntervalo3M
                        ElseIf LocalizaçãoIntervalosNoBraçoDoViolão(c, b) = "4" Then
                            CorBolinhaIntervalo = CorBolinhaIntervalo4J
                        ElseIf LocalizaçãoIntervalosNoBraçoDoViolão(c, b) = "5" Then
                            CorBolinhaIntervalo = CorBolinhaIntervalo5J
                        ElseIf LocalizaçãoIntervalosNoBraçoDoViolão(c, b) = "6" Then
                            CorBolinhaIntervalo = CorBolinhaIntervalo6M
                        ElseIf LocalizaçãoIntervalosNoBraçoDoViolão(c, b) = "7" Then
                            CorBolinhaIntervalo = CorBolinhaIntervalo7M
                        End If

                        If LocalizaçãoIntervalosNoBraçoDoViolão(c, b) = "T" OrElse LocalizaçãoIntervalosNoBraçoDoViolão(c, b) = "6" Then
                            Cor = CorFonte4
                        Else
                            Cor = CorFonte2
                        End If

                        gr.FillEllipse(CorBolinhaIntervalo, (b * 15) + 31 + AjusteLeft(0), (c * 15) + 166 + AjusteTopo, 13, 13)
                        gr.DrawEllipse(Pens.Black, (b * 15) + 31 + AjusteLeft(0), (c * 15) + 166 + AjusteTopo, 13, 13)
                        gr.DrawString(LocalizaçãoIntervalosNoBraçoDoViolão(c, b), Fonte7, Cor, (b * 15) + 38 + AjusteLeft(0) + AjusteLeftIntervalos, (c * 15) + 174 + AjusteTopo, sf)
                    End If
                Next
            Next

            AjusteLeft(0) = 29 : AjusteLeft(1) = 34 : AjusteLeft(2) = 31 : AjusteLeft(3) = 37 : AjusteLeft(4) = 39

            gr.SmoothingMode = SmoothingMode.None
            'desenha os bracinhos do violão
            'linhas horizontais
            For i = 0 To 5
                gr.DrawLine(LinhaBracinhoViolão2, 53 + AjusteLeft(0), (195 + (i * 15)) + AjusteTopo, 128 + AjusteLeft(0), (195 + (i * 15)) + AjusteTopo)
            Next
            'linhas verticais
            For i = 0 To 5
                gr.DrawLine(LinhaBracinhoViolão2, (53 + (i * 15)) + AjusteLeft(0), 179 + AjusteTopo, (53 + (i * 15)) + AjusteLeft(0), 270 + AjusteTopo)
            Next
            gr.DrawLine(Linha, 53 + AjusteLeft(0), 180 + AjusteTopo, 129 + AjusteLeft(0), 180 + AjusteTopo) 'esta linha horizontal precisa ser desenhada por último
            gr.SmoothingMode = SmoothingMode.AntiAlias


            If My.Settings.NovoValorTipoDeJogoTelaEstudoIntervalos = True Then

                Array.Clear(IntervalosDoAcordeGerado, 0, IntervalosDoAcordeGerado.Length)
                Array.Copy(IntervalosAcorde2, IntervalosAcorde, IntervalosAcorde.Length)

                ListBox1.Visible = False
                ComboBox1.Text = "" : Label1.Text = "6ª" : Label1.BackColor = Color.Transparent : Label1.ForeColor = Color.Black : ComboBox1.Visible = True
                ComboBox2.Text = "" : Label2.Text = "5ª" : Label2.BackColor = Color.Transparent : Label2.ForeColor = Color.Black : ComboBox2.Visible = True
                ComboBox3.Text = "" : Label3.Text = "4ª" : Label3.BackColor = Color.Transparent : Label3.ForeColor = Color.Black : ComboBox3.Visible = True
                ComboBox4.Text = "" : Label4.Text = "3ª" : Label4.BackColor = Color.Transparent : Label4.ForeColor = Color.Black : ComboBox4.Visible = True
                ComboBox5.Text = "" : Label5.Text = "2ª" : Label5.BackColor = Color.Transparent : Label5.ForeColor = Color.Black : ComboBox5.Visible = True
                ComboBox6.Text = "" : Label6.Text = "1ª" : Label6.BackColor = Color.Transparent : Label6.ForeColor = Color.Black : ComboBox6.Visible = True


                NúmeroAcorde = 0
                Do While NúmeroAcorde = 0
                    ' gera o array de bytes randômico de 4 bytes...
                    Dim randomNumber(3) As Byte

                    ' Create a new instance of the RNGCryptoServiceProvider.
                    Dim Gen As New Security.Cryptography.RNGCryptoServiceProvider()

                    ' Fill the array with a random value.
                    Gen.GetBytes(randomNumber)

                    ' calcula o número baseado no valor máximo
                    NúmeroAcorde = Math.Abs(BitConverter.ToInt32(randomNumber, 0)) Mod (660 + 1)
                Loop




                If ArrayAcordePadrão(NúmeroAcorde, 1, 0) = "Traste3" Then
                    AjustePosição(0) = 2
                ElseIf ArrayAcordePadrão(NúmeroAcorde, 1, 0) = "Traste4" Then
                    AjustePosição(0) = 3
                ElseIf ArrayAcordePadrão(NúmeroAcorde, 1, 0) = "Traste5" Then
                    AjustePosição(0) = 4
                ElseIf ArrayAcordePadrão(NúmeroAcorde, 1, 0) = "Traste6" Then
                    AjustePosição(0) = 5
                ElseIf ArrayAcordePadrão(NúmeroAcorde, 1, 0) = "Traste7" Then
                    AjustePosição(0) = 6
                ElseIf ArrayAcordePadrão(NúmeroAcorde, 1, 0) = "Traste8" Then
                    AjustePosição(0) = 7
                ElseIf ArrayAcordePadrão(NúmeroAcorde, 1, 0) = "Traste9" Then
                    AjustePosição(0) = 8
                ElseIf ArrayAcordePadrão(NúmeroAcorde, 1, 0) = "Traste10" Then
                    AjustePosição(0) = 9
                ElseIf ArrayAcordePadrão(NúmeroAcorde, 1, 0) = "Traste11" Then
                    AjustePosição(0) = 10
                ElseIf ArrayAcordePadrão(NúmeroAcorde, 1, 0) = "Traste12" Then
                    AjustePosição(0) = 11
                Else
                    AjustePosição(0) = 0
                End If

                AjustePosição(0) += AjustePosição(NúmeroAcorde)


                'contador += 1
                'soma += NúmeroAcorde

                If ImagemAcorde(NúmeroAcorde) = "C(#5)" OrElse ImagemAcorde(NúmeroAcorde) = "C7M(#5)" OrElse ImagemAcorde(NúmeroAcorde) = "C7M(#5)/E" OrElse _
                    ImagemAcorde(NúmeroAcorde) = "C7(#5)" OrElse ImagemAcorde(NúmeroAcorde) = "C7(#5)/E" OrElse ImagemAcorde(NúmeroAcorde) = "C7(#5)/Bb" OrElse _
                    ImagemAcorde(NúmeroAcorde) = "C7(#5/9)" OrElse ImagemAcorde(NúmeroAcorde) = "C7(#5/b9)" OrElse ImagemAcorde(NúmeroAcorde) = "C(#5) (Forma de C)" OrElse _
                    ImagemAcorde(NúmeroAcorde) = "C(#5) (Forma de A)" OrElse ImagemAcorde(NúmeroAcorde) = "C(#5) (Forma de G)" OrElse ImagemAcorde(NúmeroAcorde) = "C(#5) (Forma de E)" OrElse _
                    ImagemAcorde(NúmeroAcorde) = "C(#5) (Forma de D)" OrElse ImagemAcorde(NúmeroAcorde) = "A(#5) (Forma de A)" OrElse ImagemAcorde(NúmeroAcorde) = "A(#5) (Forma de G)" OrElse _
                    ImagemAcorde(NúmeroAcorde) = "A(#5) (Forma de E)" OrElse ImagemAcorde(NúmeroAcorde) = "A(#5) (Forma de D)" OrElse ImagemAcorde(NúmeroAcorde) = "A(#5) (Forma de C)" OrElse _
                    ImagemAcorde(NúmeroAcorde) = "G(#5) (Forma de G)" OrElse ImagemAcorde(NúmeroAcorde) = "G(#5) (Forma de E)" OrElse ImagemAcorde(NúmeroAcorde) = "G(#5) (Forma de D)" OrElse _
                    ImagemAcorde(NúmeroAcorde) = "G(#5) (Forma de C)" OrElse ImagemAcorde(NúmeroAcorde) = "G(#5) (Forma de A)" OrElse ImagemAcorde(NúmeroAcorde) = "E(#5) (Forma de E)" OrElse _
                    ImagemAcorde(NúmeroAcorde) = "E(#5) (Forma de D)" OrElse ImagemAcorde(NúmeroAcorde) = "E(#5) (Forma de C)" OrElse ImagemAcorde(NúmeroAcorde) = "E(#5) (Forma de A)" OrElse _
                    ImagemAcorde(NúmeroAcorde) = "E(#5) (Forma de G)" OrElse ImagemAcorde(NúmeroAcorde) = "D(#5) (Forma de D)" OrElse ImagemAcorde(NúmeroAcorde) = "D(#5) (Forma de C)" OrElse _
                    ImagemAcorde(NúmeroAcorde) = "D(#5) (Forma de A)" OrElse ImagemAcorde(NúmeroAcorde) = "D(#5) (Forma de G)" OrElse ImagemAcorde(NúmeroAcorde) = "D(#5) (Forma de E)" OrElse _
                    ImagemAcorde(NúmeroAcorde) = "C(#5)/E" OrElse ImagemAcorde(NúmeroAcorde) = "C(#5)/G#" OrElse ImagemAcorde(NúmeroAcorde) = "C7(#5)/G#" OrElse _
                    ImagemAcorde(NúmeroAcorde) = "C7M(#5)/G#" OrElse ImagemAcorde(NúmeroAcorde) = "C7M(#5)/B" Then
                    StringSubstituição(0) = "b6" : StringSubstituição(1) = "#5" : SubstituirString()
                ElseIf ImagemAcorde(NúmeroAcorde) = "C7M(#11)" OrElse ImagemAcorde(NúmeroAcorde) = "C7M(#11)/G" OrElse ImagemAcorde(NúmeroAcorde) = "C7M(9/#11)" OrElse _
                    ImagemAcorde(NúmeroAcorde) = "C7M(9/#11)/G" OrElse ImagemAcorde(NúmeroAcorde) = "C7(9/#11)" Then
                    StringSubstituição(0) = "b5" : StringSubstituição(1) = "#11" : SubstituirString()
                ElseIf ImagemAcorde(NúmeroAcorde) = "Cm7(11)" OrElse ImagemAcorde(NúmeroAcorde) = "Cm7(9/11)" Then
                    StringSubstituição(0) = "4" : StringSubstituição(1) = "11" : SubstituirString()
                ElseIf ImagemAcorde(NúmeroAcorde) = "C7(#11/13)" Then
                    StringSubstituição(0) = "b5" : StringSubstituição(1) = "#11" : SubstituirString()
                    StringSubstituição(0) = "6" : StringSubstituição(1) = "13" : SubstituirString()
                ElseIf ImagemAcorde(NúmeroAcorde) = "C7(13)" OrElse ImagemAcorde(NúmeroAcorde) = "C7(9/13)" OrElse ImagemAcorde(NúmeroAcorde) = "C7(b9/13)" OrElse _
                    ImagemAcorde(NúmeroAcorde) = "C7(4/13)" OrElse ImagemAcorde(NúmeroAcorde) = "C7(4/9/13)" Then
                    StringSubstituição(0) = "6" : StringSubstituição(1) = "13" : SubstituirString()
                ElseIf ImagemAcorde(NúmeroAcorde) = "C7(#9)" OrElse ImagemAcorde(NúmeroAcorde) = "C7(#9)/G" Then
                    StringSubstituição(0) = "b3" : StringSubstituição(1) = "#9" : SubstituirString()
                ElseIf ImagemAcorde(NúmeroAcorde) = "C7(#5/#9)" Then
                    StringSubstituição(0) = "b3" : StringSubstituição(1) = "#9" : SubstituirString()
                    StringSubstituição(0) = "b6" : StringSubstituição(1) = "#5" : SubstituirString()
                ElseIf ImagemAcorde(NúmeroAcorde) = "C7(#9/#11)" Then
                    StringSubstituição(0) = "b3" : StringSubstituição(1) = "#9" : SubstituirString()
                    StringSubstituição(0) = "b5" : StringSubstituição(1) = "#11" : SubstituirString()
                ElseIf ImagemAcorde(NúmeroAcorde) = "Cº" OrElse ImagemAcorde(NúmeroAcorde) = "Cº(b13)" OrElse ImagemAcorde(NúmeroAcorde) = "Cº(7M)" Then
                    StringSubstituição(0) = "6" : StringSubstituição(1) = "º" : SubstituirString()
                    StringSubstituição(0) = "b6" : StringSubstituição(1) = "b13" : SubstituirString()
                End If


                Pestana = 0
                If LinhaFinalDaSeta(NúmeroAcorde) < 6 Then
                    f = 35
                Else
                    f = 31
                End If



                CenterTextAt(gr, ImagemAcorde(NúmeroAcorde), 91 + AjusteLeft(0), 174 + AjusteTopo)
                'CenterTextAt(gr, CStr(NúmeroAcorde), 91 + AjusteLeft(0), 164 + AjusteTopo)


                Baixo = ""
                If ImagemAcorde(NúmeroAcorde).Length >= 3 Then
                    If ImagemAcorde(NúmeroAcorde).Substring(ImagemAcorde(NúmeroAcorde).Length - 2) = "/E" Then
                        If ImagemAcorde(NúmeroAcorde).Substring(0, 1) = "C" Then
                            Baixo = "Baixo na 3"
                        ElseIf ImagemAcorde(NúmeroAcorde).Substring(0, 1) = "A" Then
                            Baixo = "Baixo na 5"
                        End If
                    ElseIf ImagemAcorde(NúmeroAcorde).Substring(ImagemAcorde(NúmeroAcorde).Length - 2) = "/G" Then
                        If ImagemAcorde(NúmeroAcorde).Substring(0, 1) = "C" Then
                            Baixo = "Baixo na 5"
                        ElseIf ImagemAcorde(NúmeroAcorde).Substring(0, 1) = "A" Then
                            Baixo = "Baixo na 7"
                        End If
                    ElseIf ImagemAcorde(NúmeroAcorde).Substring(ImagemAcorde(NúmeroAcorde).Length - 3) = "/Bb" Then
                        Baixo = "Baixo na 7"
                    ElseIf ImagemAcorde(NúmeroAcorde).Substring(ImagemAcorde(NúmeroAcorde).Length - 2) = "/B" Then
                        Baixo = "Baixo na 7M"
                    ElseIf ImagemAcorde(NúmeroAcorde).Substring(ImagemAcorde(NúmeroAcorde).Length - 3) = "/Eb" Then
                        If ImagemAcorde(NúmeroAcorde).Substring(0, 1) = "C" Then
                            Baixo = "Baixo na b3"
                        ElseIf ImagemAcorde(NúmeroAcorde).Substring(0, 1) = "A" Then
                            Baixo = "Baixo na b5"
                        End If
                    ElseIf ImagemAcorde(NúmeroAcorde).Substring(ImagemAcorde(NúmeroAcorde).Length - 2) = "/F" Then
                        Baixo = "Baixo na 4"
                    ElseIf ImagemAcorde(NúmeroAcorde).Substring(ImagemAcorde(NúmeroAcorde).Length - 3) = "/G#" Then
                        If ImagemAcorde(NúmeroAcorde).Substring(0, 1) = "C" Then
                            Baixo = "Baixo na #5"
                        ElseIf ImagemAcorde(NúmeroAcorde).Substring(0, 1) = "A" Then
                            Baixo = "Baixo na 7M"
                        End If
                    ElseIf ImagemAcorde(NúmeroAcorde).Substring(ImagemAcorde(NúmeroAcorde).Length - 2) = "/A" Then
                        Baixo = "Baixo na 6"
                    ElseIf ImagemAcorde(NúmeroAcorde).Substring(ImagemAcorde(NúmeroAcorde).Length - 3) = "/Gb" Then
                        Baixo = "Baixo na b5"
                    ElseIf ImagemAcorde(NúmeroAcorde).Substring(ImagemAcorde(NúmeroAcorde).Length - 2) = "/C" Then
                        Baixo = "Baixo na b3"
                    End If
                End If

                If Baixo <> "" Then CenterTextAt(gr, Baixo, 8 + AjusteLeft(0), 225 + AjusteTopo)


                AjusteLeft(6) = 0
                d = 0
                For b = 0 To 6  'percorre as 6 colunas de cada array
                    For c = 0 To 8 'percorre as 8 linhas do array
                        If ArrayAcordePadrão(NúmeroAcorde, c, b) <> "" Then

                            If c < 7 Then
                                If ArrayAcordePadrão(NúmeroAcorde, c, b) = "1" OrElse ArrayAcordePadrão(NúmeroAcorde, c, b) = "2" OrElse ArrayAcordePadrão(NúmeroAcorde, c, b) = "3" OrElse ArrayAcordePadrão(NúmeroAcorde, c, b) = "4" OrElse _
                                   ArrayAcordePadrão(NúmeroAcorde, c, b) = "T" OrElse ArrayAcordePadrão(NúmeroAcorde, c, b) = "1T" OrElse ArrayAcordePadrão(NúmeroAcorde, c, b) = "2T" OrElse ArrayAcordePadrão(NúmeroAcorde, c, b) = "3T" OrElse ArrayAcordePadrão(NúmeroAcorde, c, b) = "4T" Then
                                    If My.Settings.NovoValorDesenharBolinhasNotasAcordes = True Then
                                        If ArrayAcordePadrão(NúmeroAcorde, c, b) <> "T" Then gr.FillEllipse(CorBolinha, (b * 15) + 31 + AjusteLeft(0), (c * 15) + 166 + AjusteTopo, 13, 13)
                                        gr.DrawString(Replace(ArrayAcordePadrão(NúmeroAcorde, c, b), "T", ""), Fonte5, CorFonte4, (b * 15) + 32 + AjusteLeft(0), (c * 15) + 166 + AjusteTopo)
                                    Else
                                        gr.DrawString(Replace(ArrayAcordePadrão(NúmeroAcorde, c, b), "T", ""), Fonte, CorFonte, (b * 15) + 32 + AjusteLeft(0), (c * 15) + 166 + AjusteTopo)
                                    End If

                                    'Desenha círculo indicando a localização da tônica de referência do acorde
                                    If My.Settings.NovoValorLocalizaçãoFundamentalDeReferenciaNoAcorde = True Then
                                        If c = 0 Then AjusteTopo += 7 'ajuste nos casos da nota de referência for uma corda solta
                                        If ArrayAcordePadrão(NúmeroAcorde, c, b) = "T" OrElse ArrayAcordePadrão(NúmeroAcorde, c, b) = "1T" OrElse ArrayAcordePadrão(NúmeroAcorde, c, b) = "2T" OrElse ArrayAcordePadrão(NúmeroAcorde, c, b) = "3T" OrElse ArrayAcordePadrão(NúmeroAcorde, c, b) = "4T" Then gr.DrawEllipse(PosiçãoNotaFundamental, (b * 15) + 31 + AjusteLeft(0), (c * 15) + 166 + AjusteTopo, 13, 13)
                                        If c = 0 Then AjusteTopo -= 7
                                    End If
                                ElseIf ArrayAcordePadrão(NúmeroAcorde, c, b) = "P" OrElse ArrayAcordePadrão(NúmeroAcorde, c, b) = "PT" Then
                                    gr.SmoothingMode = SmoothingMode.None
                                    gr.DrawLine(CorPestana, (b * 15) + 30 + AjusteLeft(1), (c * 15) + 173 + AjusteTopo, (b * 15) + f + AjusteLeft(3) + ((LinhaFinalDaSeta(NúmeroAcorde) - b) * 15), (c * 15) + 173 + AjusteTopo)
                                    gr.SmoothingMode = SmoothingMode.AntiAlias
                                    gr.DrawLine(Seta, (b * 15) + 30 + AjusteLeft(2), (c * 15) + 173 + AjusteTopo, (b * 15) + 30 + AjusteLeft(4), (c * 15) + 173 + AjusteTopo)

                                    Pestana = c

                                    'Desenha círculo indicando a localização da tônica de referência do acorde
                                    If My.Settings.NovoValorLocalizaçãoFundamentalDeReferenciaNoAcorde = True Then
                                        If ArrayAcordePadrão(NúmeroAcorde, c, b) = "PT" Then gr.DrawEllipse(PosiçãoNotaFundamental, (b * 15) + 31 + AjusteLeft(0), (c * 15) + 166 + AjusteTopo, 13, 13)
                                    End If
                                ElseIf ArrayAcordePadrão(NúmeroAcorde, c, b) = "Traste1" OrElse ArrayAcordePadrão(NúmeroAcorde, c, b) = "Traste3" OrElse ArrayAcordePadrão(NúmeroAcorde, c, b) = "Traste4" OrElse _
                                ArrayAcordePadrão(NúmeroAcorde, c, b) = "Traste5" OrElse ArrayAcordePadrão(NúmeroAcorde, c, b) = "Traste6" OrElse _
                                ArrayAcordePadrão(NúmeroAcorde, c, b) = "Traste7" OrElse ArrayAcordePadrão(NúmeroAcorde, c, b) = "Traste8" OrElse _
                                ArrayAcordePadrão(NúmeroAcorde, c, b) = "Traste9" OrElse ArrayAcordePadrão(NúmeroAcorde, c, b) = "Traste10" OrElse _
                                ArrayAcordePadrão(NúmeroAcorde, c, b) = "Traste11" OrElse ArrayAcordePadrão(NúmeroAcorde, c, b) = "Traste12" Then

                                    NumeraçãoTrastes = CInt(Replace(ArrayAcordePadrão(NúmeroAcorde, c, b), "Traste", ""))

                                    For ContadorTraste = 0 To 5
                                        AjusteLeft(5) = 0
                                        If NumeraçãoTrastes + ContadorTraste < 10 Then
                                            AjusteLeft(5) = 4
                                        End If
                                        gr.DrawString(CStr(NumeraçãoTrastes + ContadorTraste), Fonte2, CorFonte5, (b * 15) + 37 + AjusteLeft(1) + AjusteLeft(5) + AjusteLeft(6), (c * 15) + 168 + AjusteTopo + (ContadorTraste * 15))
                                    Next
                                End If
                            ElseIf c = 7 AndAlso ImagemAcorde(NúmeroAcorde) IsNot Nothing Then
                                gr.FillEllipse(Brushes.Black, (b * 15) + 34 + AjusteLeft(0), (c * 15) + 161 + AjusteTopo, 8, 8)
                                If ArrayAcordePadrão(NúmeroAcorde, 7, b) <> "" Then d += 1
                                If d = 1 Then
                                    gr.FillEllipse(Brushes.Black, (b * 15) + 35 + AjusteLeft(0), (c * 15) + 162 + AjusteTopo, 6, 6)
                                Else
                                    gr.FillEllipse(Brushes.White, (b * 15) + 35 + AjusteLeft(0), (c * 15) + 162 + AjusteTopo, 6, 6)
                                End If
                            End If
                            If c <= 6 AndAlso ArrayAcordePadrão(NúmeroAcorde, c, b) <> "" Then

                                If (ArrayAcordePadrão(NúmeroAcorde, c, b) = "P" OrElse ArrayAcordePadrão(NúmeroAcorde, c, b) = "PT") AndAlso (ArrayAcordePadrão(NúmeroAcorde, c + 1, b) = "1" OrElse ArrayAcordePadrão(NúmeroAcorde, c + 1, b) = "2" OrElse ArrayAcordePadrão(NúmeroAcorde, c + 1, b) = "3" OrElse ArrayAcordePadrão(NúmeroAcorde, c + 1, b) = "4" OrElse _
                                                                           ArrayAcordePadrão(NúmeroAcorde, c + 2, b) = "1" OrElse ArrayAcordePadrão(NúmeroAcorde, c + 2, b) = "2" OrElse ArrayAcordePadrão(NúmeroAcorde, c + 2, b) = "3" OrElse ArrayAcordePadrão(NúmeroAcorde, c + 2, b) = "4" OrElse _
                                                                           ArrayAcordePadrão(NúmeroAcorde, c + 3, b) = "1" OrElse ArrayAcordePadrão(NúmeroAcorde, c + 3, b) = "2" OrElse ArrayAcordePadrão(NúmeroAcorde, c + 3, b) = "3" OrElse ArrayAcordePadrão(NúmeroAcorde, c + 3, b) = "4" OrElse _
                                                                           ArrayAcordePadrão(NúmeroAcorde, c + 4, b) = "1" OrElse ArrayAcordePadrão(NúmeroAcorde, c + 4, b) = "2" OrElse ArrayAcordePadrão(NúmeroAcorde, c + 4, b) = "3" OrElse ArrayAcordePadrão(NúmeroAcorde, c + 4, b) = "4" OrElse _
                                                                           ArrayAcordePadrão(NúmeroAcorde, c + 5, b) = "1" OrElse ArrayAcordePadrão(NúmeroAcorde, c + 5, b) = "2" OrElse ArrayAcordePadrão(NúmeroAcorde, c + 5, b) = "3" OrElse ArrayAcordePadrão(NúmeroAcorde, c + 5, b) = "4" OrElse _
                                                                           ArrayAcordePadrão(NúmeroAcorde, c + 1, b) = "1T" OrElse ArrayAcordePadrão(NúmeroAcorde, c + 1, b) = "2T" OrElse ArrayAcordePadrão(NúmeroAcorde, c + 1, b) = "3T" OrElse ArrayAcordePadrão(NúmeroAcorde, c + 1, b) = "4T" OrElse _
                                                                           ArrayAcordePadrão(NúmeroAcorde, c + 2, b) = "1T" OrElse ArrayAcordePadrão(NúmeroAcorde, c + 2, b) = "2T" OrElse ArrayAcordePadrão(NúmeroAcorde, c + 2, b) = "3T" OrElse ArrayAcordePadrão(NúmeroAcorde, c + 2, b) = "4T" OrElse _
                                                                           ArrayAcordePadrão(NúmeroAcorde, c + 3, b) = "1T" OrElse ArrayAcordePadrão(NúmeroAcorde, c + 3, b) = "2T" OrElse ArrayAcordePadrão(NúmeroAcorde, c + 3, b) = "3T" OrElse ArrayAcordePadrão(NúmeroAcorde, c + 3, b) = "4T" OrElse _
                                                                           ArrayAcordePadrão(NúmeroAcorde, c + 4, b) = "1T" OrElse ArrayAcordePadrão(NúmeroAcorde, c + 4, b) = "2T" OrElse ArrayAcordePadrão(NúmeroAcorde, c + 4, b) = "3T" OrElse ArrayAcordePadrão(NúmeroAcorde, c + 4, b) = "4T" OrElse _
                                                                           ArrayAcordePadrão(NúmeroAcorde, c + 5, b) = "1T" OrElse ArrayAcordePadrão(NúmeroAcorde, c + 5, b) = "2T" OrElse ArrayAcordePadrão(NúmeroAcorde, c + 5, b) = "3T" OrElse ArrayAcordePadrão(NúmeroAcorde, c + 5, b) = "4T") Then
                                    'não executa nada
                                Else
                                    'gr.DrawString(IntervalosAcorde(c + AjustePosição(0), b), Fonte3, Cor, (b * 15) + 38 + AjusteLeft(0) + AjusteLeftIntervalos, 279 + AjusteTopo, sf)
                                    IntervalosDoAcordeGerado(b) = IntervalosAcorde(c + AjustePosição(0), b)
                                End If
                            End If
                        End If
                    Next


                    If ArrayAcordePadrão(NúmeroAcorde, 0, b) = "" AndAlso ArrayAcordePadrão(NúmeroAcorde, 1, b) = "" AndAlso ArrayAcordePadrão(NúmeroAcorde, 2, b) = "" AndAlso ArrayAcordePadrão(NúmeroAcorde, 3, b) = "" AndAlso ArrayAcordePadrão(NúmeroAcorde, 4, b) = "" AndAlso ArrayAcordePadrão(NúmeroAcorde, 5, b) = "" AndAlso ArrayAcordePadrão(NúmeroAcorde, 6, b) = "" Then
                        If ArrayAcordePadrão(NúmeroAcorde, 7, b) = "B" AndAlso Pestana = 0 Then
                            'gr.DrawString(IntervalosAcorde(0 + AjustePosição(1), b), Fonte3, Cor, (b * 15) + 38 + AjusteLeft(0), 279 + AjusteTopo, sf)
                            IntervalosDoAcordeGerado(b) = IntervalosAcorde(0 + AjustePosição(1), b)
                        ElseIf ArrayAcordePadrão(NúmeroAcorde, 7, b) = "B" AndAlso Pestana > 0 Then
                            'gr.DrawString(IntervalosAcorde(Pestana + AjustePosição(0), b), Fonte3, Cor, (b * 15) + 38 + AjusteLeft(0), 279 + AjusteTopo, sf)
                            IntervalosDoAcordeGerado(b) = IntervalosAcorde(Pestana + AjustePosição(0), b)
                        End If
                    End If

                Next


                If FamiliaAcorde = "Sétima e 4ª com 9ª e 13ª" OrElse NumeraçãoFamiliaAcorde = 52 Then
                    gr.DrawArc(Pens.Black, 239, 183, 25, 25, 180, 80)
                End If

                If IntervalosDoAcordeGerado(1) = "" Then ComboBox1.Visible = False : Label1.Text = "" Else gr.DrawLine(LinhaCordas, 79, 118, CInt(ComboBox1.Left + (ComboBox1.Width / 2)), 145)
                If IntervalosDoAcordeGerado(2) = "" Then ComboBox2.Visible = False : Label2.Text = "" Else gr.DrawLine(LinhaCordas, 95, 117, CInt(ComboBox2.Left + (ComboBox2.Width / 2)), 145)
                If IntervalosDoAcordeGerado(3) = "" Then ComboBox3.Visible = False : Label3.Text = "" Else gr.DrawLine(LinhaCordas, 112, 118, CInt(ComboBox3.Left + (ComboBox3.Width / 2)), 145)
                If IntervalosDoAcordeGerado(4) = "" Then ComboBox4.Visible = False : Label4.Text = "" Else gr.DrawLine(LinhaCordas, 127, 118, CInt(ComboBox4.Left + (ComboBox4.Width / 2)), 145)
                If IntervalosDoAcordeGerado(5) = "" Then ComboBox5.Visible = False : Label5.Text = "" Else gr.DrawLine(LinhaCordas, 144, 117, CInt(ComboBox5.Left + (ComboBox5.Width / 2)), 145)
                If IntervalosDoAcordeGerado(6) = "" Then ComboBox6.Visible = False : Label6.Text = "" Else gr.DrawLine(LinhaCordas, 160, 118, CInt(ComboBox6.Left + (ComboBox6.Width / 2)), 145)



            Else

                ComboBox1.Visible = False
                ComboBox2.Visible = False
                ComboBox3.Visible = False
                ComboBox4.Visible = False
                ComboBox5.Visible = False
                ComboBox6.Visible = False
                ListBox1.SelectedItem = Nothing
                ListBox1.BackColor = Color.White
                ListBox1.Visible = True

                CenterTextAt(gr, "Intervalo:", 91 + AjusteLeft(0), 174 + AjusteTopo)

                TônicaX = 0
                Do While TônicaX = 0
                    ' gera o array de bytes randômico de 4 bytes...
                    Dim randomNumber(3) As Byte

                    ' Create a new instance of the RNGCryptoServiceProvider.
                    Dim Gen As New Security.Cryptography.RNGCryptoServiceProvider()

                    ' Fill the array with a random value.
                    Gen.GetBytes(randomNumber)

                    ' calcula o número baseado no valor máximo
                    TônicaX = Math.Abs(BitConverter.ToInt32(randomNumber, 0)) Mod (6 + 1)
                Loop

                TônicaY = 0
                Do While TônicaY = 0
                    ' gera o array de bytes randômico de 4 bytes...
                    Dim randomNumber(3) As Byte

                    ' Create a new instance of the RNGCryptoServiceProvider.
                    Dim Gen As New Security.Cryptography.RNGCryptoServiceProvider()

                    ' Fill the array with a random value.
                    Gen.GetBytes(randomNumber)

                    ' calcula o número baseado no valor máximo
                    TônicaY = Math.Abs(BitConverter.ToInt32(randomNumber, 0)) Mod (6 + 1)
                Loop

                PróximaNotaX = 0
                Do While PróximaNotaX = 0 OrElse PróximaNotaX = TônicaX
                    ' gera o array de bytes randômico de 4 bytes...
                    Dim randomNumber(3) As Byte

                    ' Create a new instance of the RNGCryptoServiceProvider.
                    Dim Gen As New Security.Cryptography.RNGCryptoServiceProvider()

                    ' Fill the array with a random value.
                    Gen.GetBytes(randomNumber)

                    ' calcula o número baseado no valor máximo
                    PróximaNotaX = Math.Abs(BitConverter.ToInt32(randomNumber, 0)) Mod (6 + 1)
                Loop

                PróximaNotaY = 0
                Do While PróximaNotaY = 0
                    ' gera o array de bytes randômico de 4 bytes...
                    Dim randomNumber(3) As Byte

                    ' Create a new instance of the RNGCryptoServiceProvider.
                    Dim Gen As New Security.Cryptography.RNGCryptoServiceProvider()

                    ' Fill the array with a random value.
                    Gen.GetBytes(randomNumber)

                    ' calcula o número baseado no valor máximo
                    PróximaNotaY = Math.Abs(BitConverter.ToInt32(randomNumber, 0)) Mod (6 + 1)
                Loop

                AjusteLeft(0) = 29
                gr.FillEllipse(Brushes.White, (PróximaNotaX * 15) + 31 + AjusteLeft(0), (PróximaNotaY * 15) + 166 + AjusteTopo, 13, 13)
                gr.DrawEllipse(LinhaCordas, (PróximaNotaX * 15) + 31 + AjusteLeft(0), (PróximaNotaY * 15) + 166 + AjusteTopo, 13, 13)
                gr.DrawLine(LinhaIntervalo, (TônicaX * 15) + 38 + AjusteLeft(0), (TônicaY * 15) + 173 + AjusteTopo, (PróximaNotaX * 15) + 38 + AjusteLeft(0), (PróximaNotaY * 15) + 173 + AjusteTopo)
                gr.FillEllipse(Brushes.Black, (TônicaX * 15) + 31 + AjusteLeft(0), (TônicaY * 15) + 166 + AjusteTopo, 13, 13)


                Dim zzzz As Integer
                If TônicaX = 1 Then
                    zzzz = 6
                ElseIf TônicaX = 2 Then
                    If TônicaY <= PróximaNotaY Then
                        zzzz = 1
                    Else
                        zzzz = 13
                    End If
                ElseIf TônicaX = 3 Then
                    zzzz = 8
                ElseIf TônicaX = 4 Then
                    If TônicaY <= PróximaNotaY Then
                        zzzz = 3
                    Else
                        zzzz = 15
                    End If
                ElseIf TônicaX = 5 Then
                    zzzz = 11
                ElseIf TônicaX = 6 Then
                    zzzz = 6
                End If

                gr.DrawEllipse(RespostaIntervalo, (TônicaX * 15) + 231 + AjusteLeft(0), (zzzz * 15) + 166 + AjusteTopo, 13, 13)
                gr.DrawEllipse(RespostaIntervalo, (PróximaNotaX * 15) + 231 + AjusteLeft(0), ((zzzz + PróximaNotaY - TônicaY) * 15) + 166 + AjusteTopo, 13, 13)
                gr.DrawLine(LinhaIntervalo2, (TônicaX * 15) + 238 + AjusteLeft(0), (zzzz * 15) + 173 + AjusteTopo, (PróximaNotaX * 15) + 238 + AjusteLeft(0), ((zzzz + PróximaNotaY - TônicaY) * 15) + 173 + AjusteTopo)
                ' MsgBox(PróximaNotaY & " " & TônicaY)

                If TônicaX < PróximaNotaX Then

                    DistânciaIntervalo = (((PróximaNotaX - TônicaX) - 1) * 5) + (6 - TônicaY) + PróximaNotaY
                    If TônicaX < 5 AndAlso PróximaNotaX >= 5 Then DistânciaIntervalo -= 1

                Else

                    DistânciaIntervalo = (13 - (TônicaY - 1) - (((TônicaX - PróximaNotaX) - 1) * 5)) - (6 - PróximaNotaY)
                    If TônicaX >= 5 AndAlso PróximaNotaX < 5 Then DistânciaIntervalo += 1

                End If

                Do While DistânciaIntervalo > 12

                    DistânciaIntervalo -= 12

                Loop

                Do While DistânciaIntervalo < 1

                    DistânciaIntervalo += 12

                Loop

                NomeIntervalo = ""
                If DistânciaIntervalo = 1 Then
                    NomeIntervalo = "T"
                ElseIf DistânciaIntervalo = 2 Then
                    NomeIntervalo = "b9"
                ElseIf DistânciaIntervalo = 3 Then
                    NomeIntervalo = "2 / 9"
                ElseIf DistânciaIntervalo = 4 Then
                    NomeIntervalo = "b3 / #9"
                ElseIf DistânciaIntervalo = 5 Then
                    NomeIntervalo = "3"
                ElseIf DistânciaIntervalo = 6 Then
                    NomeIntervalo = "4 / 11"
                ElseIf DistânciaIntervalo = 7 Then
                    NomeIntervalo = "b5 / #11"
                ElseIf DistânciaIntervalo = 8 Then
                    NomeIntervalo = "5"
                ElseIf DistânciaIntervalo = 9 Then
                    NomeIntervalo = "#5 / b6 / b13"
                ElseIf DistânciaIntervalo = 10 Then
                    NomeIntervalo = "6 / 13 / º"
                ElseIf DistânciaIntervalo = 11 Then
                    NomeIntervalo = "7"
                ElseIf DistânciaIntervalo = 12 Then
                    NomeIntervalo = "7M"
                End If


                'CenterTextAt(gr, NomeIntervalo, 91 + AjusteLeft(0), 164 + AjusteTopo)
            End If

            gr.SmoothingMode = SmoothingMode.None
            'bmp.SetResolution(96 * FatorEscala, 96 * FatorEscala)
            'Clipboard.SetDataObject(bmp)
            Me.BackgroundImage = bmp

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Public Sub GerarAcordes()

        Try

            LocalizaçãoIntervalosNoBraçoDoViolão(1, 1) = "5"
            LocalizaçãoIntervalosNoBraçoDoViolão(1, 2) = "T"
            LocalizaçãoIntervalosNoBraçoDoViolão(1, 3) = "4"
            LocalizaçãoIntervalosNoBraçoDoViolão(1, 5) = "2"
            LocalizaçãoIntervalosNoBraçoDoViolão(1, 6) = "5"

            LocalizaçãoIntervalosNoBraçoDoViolão(2, 4) = "7"

            LocalizaçãoIntervalosNoBraçoDoViolão(3, 1) = "6"
            LocalizaçãoIntervalosNoBraçoDoViolão(3, 2) = "2"
            LocalizaçãoIntervalosNoBraçoDoViolão(3, 3) = "5"
            LocalizaçãoIntervalosNoBraçoDoViolão(3, 4) = "T"
            LocalizaçãoIntervalosNoBraçoDoViolão(3, 5) = "3"
            LocalizaçãoIntervalosNoBraçoDoViolão(3, 6) = "6"

            LocalizaçãoIntervalosNoBraçoDoViolão(4, 5) = "4"

            LocalizaçãoIntervalosNoBraçoDoViolão(5, 1) = "7"
            LocalizaçãoIntervalosNoBraçoDoViolão(5, 2) = "3"
            LocalizaçãoIntervalosNoBraçoDoViolão(5, 3) = "6"
            LocalizaçãoIntervalosNoBraçoDoViolão(5, 4) = "2"
            LocalizaçãoIntervalosNoBraçoDoViolão(5, 6) = "7"

            LocalizaçãoIntervalosNoBraçoDoViolão(6, 1) = "T"
            LocalizaçãoIntervalosNoBraçoDoViolão(6, 2) = "4"
            LocalizaçãoIntervalosNoBraçoDoViolão(6, 5) = "5"
            LocalizaçãoIntervalosNoBraçoDoViolão(6, 6) = "T"

            LocalizaçãoIntervalosNoBraçoDoViolão(7, 3) = "7"
            LocalizaçãoIntervalosNoBraçoDoViolão(7, 4) = "3"

            LocalizaçãoIntervalosNoBraçoDoViolão(8, 1) = "2"
            LocalizaçãoIntervalosNoBraçoDoViolão(8, 2) = "5"
            LocalizaçãoIntervalosNoBraçoDoViolão(8, 3) = "T"
            LocalizaçãoIntervalosNoBraçoDoViolão(8, 4) = "4"
            LocalizaçãoIntervalosNoBraçoDoViolão(8, 5) = "6"
            LocalizaçãoIntervalosNoBraçoDoViolão(8, 6) = "2"

            LocalizaçãoIntervalosNoBraçoDoViolão(10, 1) = "3"
            LocalizaçãoIntervalosNoBraçoDoViolão(10, 2) = "6"
            LocalizaçãoIntervalosNoBraçoDoViolão(10, 3) = "2"
            LocalizaçãoIntervalosNoBraçoDoViolão(10, 4) = "5"
            LocalizaçãoIntervalosNoBraçoDoViolão(10, 5) = "7"
            LocalizaçãoIntervalosNoBraçoDoViolão(10, 6) = "3"

            LocalizaçãoIntervalosNoBraçoDoViolão(11, 1) = "4"
            LocalizaçãoIntervalosNoBraçoDoViolão(11, 5) = "T"
            LocalizaçãoIntervalosNoBraçoDoViolão(11, 6) = "4"

            LocalizaçãoIntervalosNoBraçoDoViolão(12, 2) = "7"
            LocalizaçãoIntervalosNoBraçoDoViolão(12, 3) = "3"
            LocalizaçãoIntervalosNoBraçoDoViolão(12, 4) = "6"

            LocalizaçãoIntervalosNoBraçoDoViolão(13, 1) = "5"
            LocalizaçãoIntervalosNoBraçoDoViolão(13, 2) = "T"
            LocalizaçãoIntervalosNoBraçoDoViolão(13, 3) = "4"
            LocalizaçãoIntervalosNoBraçoDoViolão(13, 5) = "2"
            LocalizaçãoIntervalosNoBraçoDoViolão(13, 6) = "5"

            LocalizaçãoIntervalosNoBraçoDoViolão(14, 4) = "7"

            LocalizaçãoIntervalosNoBraçoDoViolão(15, 1) = "6"
            LocalizaçãoIntervalosNoBraçoDoViolão(15, 2) = "2"
            LocalizaçãoIntervalosNoBraçoDoViolão(15, 3) = "5"
            LocalizaçãoIntervalosNoBraçoDoViolão(15, 4) = "T"
            LocalizaçãoIntervalosNoBraçoDoViolão(15, 5) = "3"
            LocalizaçãoIntervalosNoBraçoDoViolão(15, 6) = "6"

            LocalizaçãoIntervalosNoBraçoDoViolão(16, 5) = "4"

            LocalizaçãoIntervalosNoBraçoDoViolão(17, 1) = "7"
            LocalizaçãoIntervalosNoBraçoDoViolão(17, 2) = "3"
            LocalizaçãoIntervalosNoBraçoDoViolão(17, 3) = "6"
            LocalizaçãoIntervalosNoBraçoDoViolão(17, 4) = "2"
            LocalizaçãoIntervalosNoBraçoDoViolão(17, 6) = "7"

            LocalizaçãoIntervalosNoBraçoDoViolão(18, 1) = "T"
            LocalizaçãoIntervalosNoBraçoDoViolão(18, 2) = "4"
            LocalizaçãoIntervalosNoBraçoDoViolão(18, 5) = "5"
            LocalizaçãoIntervalosNoBraçoDoViolão(18, 6) = "T"

            LocalizaçãoIntervalosNoBraçoDoViolão(19, 3) = "7"
            LocalizaçãoIntervalosNoBraçoDoViolão(19, 4) = "3"

            LocalizaçãoIntervalosNoBraçoDoViolão(20, 1) = "2"
            LocalizaçãoIntervalosNoBraçoDoViolão(20, 2) = "5"
            LocalizaçãoIntervalosNoBraçoDoViolão(20, 3) = "T"
            LocalizaçãoIntervalosNoBraçoDoViolão(20, 4) = "4"
            LocalizaçãoIntervalosNoBraçoDoViolão(20, 5) = "6"
            LocalizaçãoIntervalosNoBraçoDoViolão(20, 6) = "2"

            'Array.Clear(ArrayAcordePadrão, 0, ArrayAcordePadrão.Length)
            'Array.Clear(ImagemAcorde, 0, ImagemAcorde.Length)
            'Array.Clear(LinhaFinalDaSeta, 0, LinhaFinalDaSeta.Length)

            'melhor definir isso para todos os campos da linha 6 do array logo no começo, e daí nos acordes, limpa-se os campos que não terão valor. Isso gera menos código
            For i = 1 To 660
                ArrayAcordePadrão(i, 7, 1) = "B" : ArrayAcordePadrão(i, 7, 2) = "B" : ArrayAcordePadrão(i, 7, 3) = "B" : ArrayAcordePadrão(i, 7, 4) = "B" : ArrayAcordePadrão(i, 7, 5) = "B" : ArrayAcordePadrão(i, 7, 6) = "B"
                ArrayAcordePadrão(i, 1, 0) = "Traste1"
            Next


            Index = 0
            'Array.Clear(AjustePosição, 0, AjustePosição.Length)

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

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3T"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*" 'indica se este acorde está entre os mais usados 
            'ArrayAcordePadrão(Index,8, 2) = "T" : ArrayAcordePadrão(Index,8, 3) = "3" : ArrayAcordePadrão(Index,8, 4) = "5" : ArrayAcordePadrão(Index,8, 5) = "T" : ArrayAcordePadrão(Index,8, 6) = "3"
            ImagemAcorde(Index) = "C"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 5, 3) = "2"
            ArrayAcordePadrão(Index, 5, 4) = "3"
            ArrayAcordePadrão(Index, 5, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 2) = "2"
            ArrayAcordePadrão(Index, 4, 1) = "3T"
            ArrayAcordePadrão(Index, 4, 6) = "4"
            ImagemAcorde(Index) = "C"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 5
            ArrayAcordePadrão(Index, 4, 1) = "4T"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 3, 3) = "4"
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8" 'indica a casa inicial do acorde 

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 3, 4) = "2"
            ArrayAcordePadrão(Index, 3, 6) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1T"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C/E"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1T"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ImagemAcorde(Index) = "C/E"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2"
            ArrayAcordePadrão(Index, 4, 5) = "3"
            ArrayAcordePadrão(Index, 4, 6) = "4T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 4) = "T"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 3, 1) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4T"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 3, 1) = "3"
            ArrayAcordePadrão(Index, 3, 2) = "4T"
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C/G"

            Index += 1
            ArrayAcordePadrão(Index, 3, 1) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 3, 2) = "T"
            ArrayAcordePadrão(Index, 5, 3) = "2"
            ArrayAcordePadrão(Index, 5, 4) = "3"
            ArrayAcordePadrão(Index, 5, 5) = "4"
            ImagemAcorde(Index) = "C/G"

            Index += 1
            ArrayAcordePadrão(Index, 3, 6) = "1"
            ArrayAcordePadrão(Index, 5, 3) = "2"
            ArrayAcordePadrão(Index, 5, 4) = "3T"
            ArrayAcordePadrão(Index, 5, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C/G"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 6) = "T"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 3, 3) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "1"
            ArrayAcordePadrão(Index, 3, 4) = "2"
            ArrayAcordePadrão(Index, 3, 6) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ImagemAcorde(Index) = "C/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3T"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ImagemAcorde(Index) = "C(add9)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 3) = "2"
            ArrayAcordePadrão(Index, 3, 5) = "3"
            ArrayAcordePadrão(Index, 5, 4) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C(add9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste3"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 2) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 4, 1) = "4T"
            ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C(add9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 5, 3) = "4"
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C(add9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 3) = "3T"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C(add9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 5) = "2T"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C(add9)/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste12"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C(#5)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 4) = "2"
            ArrayAcordePadrão(Index, 3, 5) = "3"
            ArrayAcordePadrão(Index, 4, 3) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C(#5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste3"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 4, 1) = "4T"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C(#5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 2, 5) = "3"
            ArrayAcordePadrão(Index, 3, 3) = "4"
            ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C(#5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 3) = "3"
            ArrayAcordePadrão(Index, 4, 2) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C(#5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 6) = "2"
            ArrayAcordePadrão(Index, 4, 4) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C(#5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 2, 4) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 3, 6) = "3"
            ArrayAcordePadrão(Index, 5, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C6"

            Index += 1
            ArrayAcordePadrão(Index, 2, 4) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 5, 3) = "3"
            ArrayAcordePadrão(Index, 5, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C6"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 5, 3) = "4"
            ArrayAcordePadrão(Index, 5, 4) = "4"
            ArrayAcordePadrão(Index, 5, 5) = "4"
            ArrayAcordePadrão(Index, 5, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ImagemAcorde(Index) = "C6"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 4) = "2"
            ArrayAcordePadrão(Index, 3, 5) = "3"
            ArrayAcordePadrão(Index, 5, 3) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ImagemAcorde(Index) = "C6"
            ArrayAcordePadrão(Index, 1, 0) = "Traste3"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 2, 1) = "2T"
            ArrayAcordePadrão(Index, 2, 5) = "3"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C6"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 2, 1) = "2T"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 4, 2) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C6"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ImagemAcorde(Index) = "C6"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 6) = "1"
            ArrayAcordePadrão(Index, 4, 1) = "2T"
            ArrayAcordePadrão(Index, 4, 5) = "3"
            ArrayAcordePadrão(Index, 5, 4) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ImagemAcorde(Index) = "C6"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C6"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "1"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 1) = "3"
            ArrayAcordePadrão(Index, 3, 2) = "4T"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C6/G"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 4) = "T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C6/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 2, 2) = "2"
            ArrayAcordePadrão(Index, 2, 3) = "3T"
            ArrayAcordePadrão(Index, 2, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C6/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste9"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "1"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3T"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C6(9)"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "P" : LinhaFinalDaSeta(Index) = 5
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 3, 5) = "3"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C6(9)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "P" : LinhaFinalDaSeta(Index) = 4
            ArrayAcordePadrão(Index, 2, 1) = "2T"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C6(9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 1) = "2T"
            ArrayAcordePadrão(Index, 2, 5) = "3"
            ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C6(9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 1) = "2T"
            ArrayAcordePadrão(Index, 2, 5) = "3"
            ArrayAcordePadrão(Index, 2, 6) = "4"
            ImagemAcorde(Index) = "C6(9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 2, 3) = "2T"
            ArrayAcordePadrão(Index, 2, 5) = "3"
            ArrayAcordePadrão(Index, 2, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C6(9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste9"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 5) = "2"
            ArrayAcordePadrão(Index, 2, 6) = "3T"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ImagemAcorde(Index) = "C6(9)/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "P" : LinhaFinalDaSeta(Index) = 5
            ArrayAcordePadrão(Index, 3, 1) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3T"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C6(9)/G"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 2, 2) = "2"
            ArrayAcordePadrão(Index, 2, 3) = "2T"
            ArrayAcordePadrão(Index, 2, 5) = "3"
            ArrayAcordePadrão(Index, 2, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ImagemAcorde(Index) = "C6(9)/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste9"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7M"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 4, 4) = "2"
            ArrayAcordePadrão(Index, 5, 3) = "3"
            ArrayAcordePadrão(Index, 5, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7M"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "1T"
            ArrayAcordePadrão(Index, 3, 4) = "2"
            ArrayAcordePadrão(Index, 3, 5) = "3"
            ArrayAcordePadrão(Index, 5, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ImagemAcorde(Index) = "C7M"
            ArrayAcordePadrão(Index, 1, 0) = "Traste3"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 2, 4) = "3"
            ArrayAcordePadrão(Index, 3, 2) = "4"
            ImagemAcorde(Index) = "C7M"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "1T"
            ArrayAcordePadrão(Index, 1, 5) = "2"
            ArrayAcordePadrão(Index, 2, 3) = "3"
            ArrayAcordePadrão(Index, 2, 4) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7M"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 6) = "1"
            ArrayAcordePadrão(Index, 2, 1) = "2T"
            ArrayAcordePadrão(Index, 2, 5) = "3"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7M"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 3) = "3"
            ArrayAcordePadrão(Index, 5, 5) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 3, 4) = "2"
            ArrayAcordePadrão(Index, 3, 5) = "3"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7M"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 6) = "1"
            ArrayAcordePadrão(Index, 2, 5) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 4, 3) = "4T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7M"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1T"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 3, 6) = "3"
            ArrayAcordePadrão(Index, 4, 4) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7M/E"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1T"
            ArrayAcordePadrão(Index, 3, 2) = "2"
            ArrayAcordePadrão(Index, 3, 6) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7M/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1T"
            ArrayAcordePadrão(Index, 3, 2) = "2"
            ArrayAcordePadrão(Index, 4, 5) = "3"
            ArrayAcordePadrão(Index, 5, 3) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 3, 1) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7M/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 3, 1) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 5, 2) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 3, 1) = "3"
            ArrayAcordePadrão(Index, 3, 2) = "4T"
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7M/G"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 4) = "T"
            ArrayAcordePadrão(Index, 3, 6) = "3"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7M/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3T"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7M(#5)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "1T"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 5) = "3"
            ArrayAcordePadrão(Index, 4, 3) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M(#5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste3"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "1T"
            ArrayAcordePadrão(Index, 4, 4) = "2"
            ArrayAcordePadrão(Index, 4, 6) = "3"
            ArrayAcordePadrão(Index, 5, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ImagemAcorde(Index) = "C7M(#5)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "1T"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 2, 4) = "3"
            ArrayAcordePadrão(Index, 2, 5) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7M(#5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 3, 5) = "2"
            ArrayAcordePadrão(Index, 3, 6) = "3"
            ArrayAcordePadrão(Index, 4, 4) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7M(#5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 3, 1) = "2"
            ArrayAcordePadrão(Index, 3, 5) = "3"
            ArrayAcordePadrão(Index, 4, 4) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M(#5)/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "1"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3T"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7M(6)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "1T"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 5) = "3"
            ArrayAcordePadrão(Index, 5, 3) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M(6)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste3"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "1T"
            ArrayAcordePadrão(Index, 4, 4) = "2"
            ArrayAcordePadrão(Index, 5, 5) = "3"
            ArrayAcordePadrão(Index, 5, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7M(6)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "1T"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 2, 4) = "3"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7M(6)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 3, 5) = "2"
            ArrayAcordePadrão(Index, 3, 6) = "3"
            ArrayAcordePadrão(Index, 5, 4) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7M(6)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 3, 5) = "3"
            ArrayAcordePadrão(Index, 4, 4) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7M(9)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "P" : LinhaFinalDaSeta(Index) = 4
            ArrayAcordePadrão(Index, 2, 1) = "2T"
            ArrayAcordePadrão(Index, 3, 3) = "3"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7M(9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 6) = "3"
            ArrayAcordePadrão(Index, 5, 5) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7M(9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 1) = "2T"
            ArrayAcordePadrão(Index, 2, 5) = "3"
            ArrayAcordePadrão(Index, 4, 3) = "4"
            ImagemAcorde(Index) = "C7M(9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 2, 4) = "3"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7M(9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 2, 3) = "2T"
            ArrayAcordePadrão(Index, 2, 6) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7M(9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste9"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 5) = "2"
            ArrayAcordePadrão(Index, 4, 3) = "4T"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ImagemAcorde(Index) = "C7M(9)/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "1"
            ArrayAcordePadrão(Index, 3, 1) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 3, 5) = "3"
            ArrayAcordePadrão(Index, 4, 4) = "4"
            ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M(9)/G"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 2, 2) = "2"
            ArrayAcordePadrão(Index, 2, 3) = "2T"
            ArrayAcordePadrão(Index, 2, 6) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ImagemAcorde(Index) = "C7M(9)/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste9"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 5) = "2"
            ArrayAcordePadrão(Index, 4, 3) = "4T"
            ImagemAcorde(Index) = "C7M(9)/B"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 1) = "2T"
            ArrayAcordePadrão(Index, 2, 5) = "3"
            ImagemAcorde(Index) = "C7M(6/9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "1"
            ArrayAcordePadrão(Index, 2, 6) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3T"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ImagemAcorde(Index) = "C7M(#11)"

            Index += 1
            ArrayAcordePadrão(Index, 2, 6) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 4, 4) = "3"
            ArrayAcordePadrão(Index, 5, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7M(#11)"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "1T"
            ArrayAcordePadrão(Index, 4, 3) = "2"
            ArrayAcordePadrão(Index, 4, 4) = "3"
            ArrayAcordePadrão(Index, 5, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M(#11)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1"
            ArrayAcordePadrão(Index, 2, 1) = "2T"
            ArrayAcordePadrão(Index, 3, 3) = "3"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7M(#11)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "1T"
            ArrayAcordePadrão(Index, 2, 2) = "2"
            ArrayAcordePadrão(Index, 2, 3) = "3"
            ArrayAcordePadrão(Index, 2, 4) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M(#11)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 1) = "2T"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ImagemAcorde(Index) = "C7M(#11)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 1) = "2T"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 4, 3) = "4"
            ImagemAcorde(Index) = "C7M(#11)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 4, 3) = "4T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7M(#11)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "1"
            ArrayAcordePadrão(Index, 2, 6) = "2"
            ArrayAcordePadrão(Index, 3, 1) = "3"
            ArrayAcordePadrão(Index, 3, 2) = "4T"
            ImagemAcorde(Index) = "C7M(#11)/G"

            Index += 1
            ArrayAcordePadrão(Index, 2, 2) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 3, 5) = "3"
            ArrayAcordePadrão(Index, 4, 4) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ImagemAcorde(Index) = "C7M(9/#11)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "P" : LinhaFinalDaSeta(Index) = 5
            ArrayAcordePadrão(Index, 2, 1) = "2T"
            ArrayAcordePadrão(Index, 3, 3) = "3"
            ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M(9/#11)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 1) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 3, 5) = "3"
            ArrayAcordePadrão(Index, 4, 4) = "4"
            ImagemAcorde(Index) = "C7M(9/#11)/G"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 1, 5) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "4T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 4, 5) = "2"
            ArrayAcordePadrão(Index, 5, 3) = "3"
            ArrayAcordePadrão(Index, 5, 4) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 3, 6) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ImagemAcorde(Index) = "Cm"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 4, 1) = "2T"
            ArrayAcordePadrão(Index, 4, 4) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 3, 3) = "4"
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 2, 6) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 1, 5) = "2T"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "Cm/Eb"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 2, 2) = "2"
            ArrayAcordePadrão(Index, 4, 5) = "3"
            ArrayAcordePadrão(Index, 4, 6) = "4T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm/Eb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 5 : ArrayAcordePadrão(Index, 1, 4) = "T"
            ArrayAcordePadrão(Index, 2, 2) = "2"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm/Eb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 2, 1) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4T"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm/Eb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 6) = "1"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 3) = "3"
            ArrayAcordePadrão(Index, 3, 5) = "4T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm/Eb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste11"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 5) = "3T"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm/Eb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste11"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 1, 5) = "2T"
            ArrayAcordePadrão(Index, 3, 1) = "3"
            ArrayAcordePadrão(Index, 3, 2) = "4"
            ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm/G"

            Index += 1
            ArrayAcordePadrão(Index, 3, 1) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 3, 2) = "T"
            ArrayAcordePadrão(Index, 4, 5) = "2"
            ArrayAcordePadrão(Index, 5, 3) = "3"
            ArrayAcordePadrão(Index, 5, 4) = "4"
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm/G"

            Index += 1
            ArrayAcordePadrão(Index, 3, 6) = "1"
            ArrayAcordePadrão(Index, 4, 5) = "2"
            ArrayAcordePadrão(Index, 5, 3) = "3"
            ArrayAcordePadrão(Index, 5, 4) = "4T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm/G"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 6) = "T"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 3, 3) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 3) = "T"
            ArrayAcordePadrão(Index, 2, 6) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ImagemAcorde(Index) = "Cm/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "3T"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm(add9)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 5) = "2"
            ArrayAcordePadrão(Index, 3, 3) = "3"
            ArrayAcordePadrão(Index, 5, 4) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm(add9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste3"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 2, 2) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 4, 1) = "4T"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm(add9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 5, 3) = "4"
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm(add9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 3) = "3T"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm(add9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 2, 4) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 3, 6) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm6"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm6"

            Index += 1
            ArrayAcordePadrão(Index, 2, 4) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 4, 5) = "3"
            ArrayAcordePadrão(Index, 5, 3) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm6"

            Index += 1
            ArrayAcordePadrão(Index, 2, 4) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 4, 5) = "3"
            ArrayAcordePadrão(Index, 5, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ImagemAcorde(Index) = "Cm6"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 2, 1) = "2T"
            ArrayAcordePadrão(Index, 2, 4) = "3"
            ArrayAcordePadrão(Index, 2, 5) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm6"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 2, 1) = "2T"
            ArrayAcordePadrão(Index, 2, 4) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm6"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 6) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "Cm6"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 4) = "T"
            ArrayAcordePadrão(Index, 2, 2) = "2"
            ArrayAcordePadrão(Index, 3, 3) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm6/Eb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 1) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm6/Eb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 2, 4) = "1"
            ArrayAcordePadrão(Index, 3, 1) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 3, 6) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 3) = ""
            ImagemAcorde(Index) = "Cm6/G"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 2, 4) = "3T"
            ArrayAcordePadrão(Index, 2, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "Cm6/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste4"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2"
            ArrayAcordePadrão(Index, 3, 3) = "3T"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm6/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3T"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm6(9)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 2, 1) = "2T"
            ArrayAcordePadrão(Index, 2, 4) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 4, 6) = "4"
            ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "Cm6(9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 3, 3) = "2T"
            ArrayAcordePadrão(Index, 3, 5) = "3"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm6(9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 3) = "T"
            ArrayAcordePadrão(Index, 2, 1) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm6(9)/Eb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 2) = "3T"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm7"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 4, 5) = "2"
            ArrayAcordePadrão(Index, 5, 3) = "3"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm7"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "1T"
            ArrayAcordePadrão(Index, 3, 4) = "2"
            ArrayAcordePadrão(Index, 3, 6) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 5) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 4, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ImagemAcorde(Index) = "Cm7"
            ArrayAcordePadrão(Index, 1, 0) = "Traste3"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm7"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "1T"
            ArrayAcordePadrão(Index, 1, 3) = "2"
            ArrayAcordePadrão(Index, 1, 4) = "3"
            ArrayAcordePadrão(Index, 1, 5) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm7"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 2) = "2"
            ArrayAcordePadrão(Index, 3, 3) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm7"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 2, 5) = "2"
            ArrayAcordePadrão(Index, 2, 6) = "3"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm7"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 5) = "T"
            ArrayAcordePadrão(Index, 3, 1) = "3"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm7/G"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 6) = "T"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ImagemAcorde(Index) = "Cm7/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 5) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 4, 1) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm/Bb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste3"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 2, 4) = "3T"
            ArrayAcordePadrão(Index, 3, 1) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm/Bb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste4"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "1"
            ArrayAcordePadrão(Index, 3, 4) = "2"
            ArrayAcordePadrão(Index, 3, 5) = "3"
            ArrayAcordePadrão(Index, 3, 6) = "4T"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm/Bb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste6"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 6) = "T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm/Bb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 6) = "1"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 3, 5) = "4T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm/Bb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste11"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm7(9)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ImagemAcorde(Index) = "Cm7(9)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "1"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 1) = "3T"
            ArrayAcordePadrão(Index, 3, 3) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm7(9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste6"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 2) = "2"
            ArrayAcordePadrão(Index, 3, 6) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm7(9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ImagemAcorde(Index) = "Cm7(9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 3, 3) = "2T"
            ArrayAcordePadrão(Index, 3, 6) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm7(9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 3, 1) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm7(9)/G"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "1T"
            ArrayAcordePadrão(Index, 3, 4) = "2"
            ArrayAcordePadrão(Index, 4, 3) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm7(b5)"

            Index += 1
            ArrayAcordePadrão(Index, 2, 6) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm7(b5)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 2, 5) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 4, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ImagemAcorde(Index) = "Cm7(b5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste3"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1"
            ArrayAcordePadrão(Index, 2, 1) = "2T"
            ArrayAcordePadrão(Index, 2, 3) = "3"
            ArrayAcordePadrão(Index, 2, 4) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm7(b5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 2) = "2"
            ArrayAcordePadrão(Index, 3, 3) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm7(b5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "1T"
            ArrayAcordePadrão(Index, 1, 3) = "2"
            ArrayAcordePadrão(Index, 1, 4) = "3"
            ArrayAcordePadrão(Index, 2, 2) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm7(b5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 2, 5) = "3"
            ArrayAcordePadrão(Index, 2, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm7(b5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 2) = "3T"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ImagemAcorde(Index) = "Cm7(11)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 6) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm7(11)"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 4, 5) = "2"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm7(11)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 5) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 4, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm7(11)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste3"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1"
            ArrayAcordePadrão(Index, 3, 1) = "2T"
            ArrayAcordePadrão(Index, 3, 3) = "3"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm7(11)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste6"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 3) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm7(11)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "1T"
            ArrayAcordePadrão(Index, 1, 2) = "2"
            ArrayAcordePadrão(Index, 1, 3) = "3"
            ArrayAcordePadrão(Index, 1, 4) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm7(11)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 1, 4) = "2"
            ArrayAcordePadrão(Index, 2, 5) = "3"
            ArrayAcordePadrão(Index, 2, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm7(11)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm7(9/11)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 1) = "3T"
            ArrayAcordePadrão(Index, 3, 3) = "4"
            ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm7(9/11)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste6"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "3T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm(7M)"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 4, 4) = "2"
            ArrayAcordePadrão(Index, 4, 5) = "3"
            ArrayAcordePadrão(Index, 5, 3) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm(7M)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm(7M)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 2, 6) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm(7M)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 6) = "1"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 2, 5) = "3"
            ArrayAcordePadrão(Index, 4, 3) = "4T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "Cm(7M)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm(7M/6)"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 4, 4) = "2"
            ArrayAcordePadrão(Index, 4, 5) = "3"
            ArrayAcordePadrão(Index, 5, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ImagemAcorde(Index) = "Cm(7M/6)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 1) = "2T"
            ArrayAcordePadrão(Index, 2, 4) = "3"
            ArrayAcordePadrão(Index, 2, 5) = "4"
            ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm(7M/6)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 3, 5) = "3"
            ArrayAcordePadrão(Index, 4, 4) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm(7M/9)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "1"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 1) = "3T"
            ArrayAcordePadrão(Index, 4, 3) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm(7M/9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste6"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ImagemAcorde(Index) = "Cm(7M/9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 6) = "3"
            ArrayAcordePadrão(Index, 5, 5) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cm(7M/9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 2) = "3T"
            ArrayAcordePadrão(Index, 3, 3) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ImagemAcorde(Index) = "C4"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 3) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C4"
            ArrayAcordePadrão(Index, 1, 0) = "Traste3"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 2) = "2"
            ArrayAcordePadrão(Index, 3, 3) = "3"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C4"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 3, 4) = "2"
            ArrayAcordePadrão(Index, 4, 5) = "3"
            ArrayAcordePadrão(Index, 4, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C4"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3T"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 5, 3) = "3"
            ArrayAcordePadrão(Index, 5, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 4) = "2"
            ArrayAcordePadrão(Index, 3, 5) = "3"
            ArrayAcordePadrão(Index, 4, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7"
            ArrayAcordePadrão(Index, 1, 0) = "Traste3"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 3) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "2"
            ArrayAcordePadrão(Index, 3, 5) = "3"
            ArrayAcordePadrão(Index, 4, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ImagemAcorde(Index) = "C7"
            ArrayAcordePadrão(Index, 1, 0) = "Traste3"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 6) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 4, 1) = "4T"
            ImagemAcorde(Index) = "C7"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "1T"
            ArrayAcordePadrão(Index, 1, 3) = "2"
            ArrayAcordePadrão(Index, 1, 5) = "3"
            ArrayAcordePadrão(Index, 2, 4) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 3) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 2, 5) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1T"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7/E"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1T"
            ArrayAcordePadrão(Index, 2, 6) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1T"
            ArrayAcordePadrão(Index, 3, 2) = "2"
            ArrayAcordePadrão(Index, 4, 3) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 4) = "T"
            ArrayAcordePadrão(Index, 2, 6) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 5) = ""
            ImagemAcorde(Index) = "C7/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 2, 5) = "2"
            ArrayAcordePadrão(Index, 3, 1) = "3"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1T"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 3, 1) = "3"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7/G"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "1"
            ArrayAcordePadrão(Index, 3, 4) = "2T"
            ArrayAcordePadrão(Index, 3, 5) = "3"
            ArrayAcordePadrão(Index, 4, 6) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ImagemAcorde(Index) = "C7/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste3"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 4) = "T"
            ArrayAcordePadrão(Index, 2, 6) = "2"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 2, 2) = "2"
            ArrayAcordePadrão(Index, 2, 3) = "3T"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste9"

            Index += 1
            ArrayAcordePadrão(Index, 1, 6) = "1T"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "2"
            ArrayAcordePadrão(Index, 1, 5) = "3T"
            ArrayAcordePadrão(Index, 2, 3) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C/Bb"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "1"
            ArrayAcordePadrão(Index, 1, 5) = "2T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C/Bb"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 4) = "T"
            ArrayAcordePadrão(Index, 2, 1) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C/Bb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "1"
            ArrayAcordePadrão(Index, 3, 5) = "2"
            ArrayAcordePadrão(Index, 3, 6) = "3T"
            ArrayAcordePadrão(Index, 4, 4) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ImagemAcorde(Index) = "C/Bb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste6"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 4) = "T"
            ArrayAcordePadrão(Index, 2, 1) = "2"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C/Bb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 6) = "T"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C/Bb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 2, 6) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 5, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(b5)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "1T"
            ArrayAcordePadrão(Index, 1, 4) = "2"
            ArrayAcordePadrão(Index, 2, 3) = "3"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(b5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste3"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1"
            ArrayAcordePadrão(Index, 2, 1) = "2T"
            ArrayAcordePadrão(Index, 2, 3) = "3"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(b5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "1T"
            ArrayAcordePadrão(Index, 1, 3) = "2"
            ArrayAcordePadrão(Index, 2, 2) = "3"
            ArrayAcordePadrão(Index, 2, 4) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(b5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 2, 5) = "3"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7(b5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1T"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 2, 6) = "3"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7(b5)/E"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(9/#11)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 1) = "2T"
            ArrayAcordePadrão(Index, 2, 3) = "3"
            ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(9/#11)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 6) = "1"
            ArrayAcordePadrão(Index, 3, 5) = "2"
            ArrayAcordePadrão(Index, 4, 1) = "3T"
            ArrayAcordePadrão(Index, 4, 3) = "3"
            ArrayAcordePadrão(Index, 5, 4) = "4"
            ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7(#11/13)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 4, 6) = "2"
            ArrayAcordePadrão(Index, 5, 5) = "3"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(#5)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "1T"
            ArrayAcordePadrão(Index, 1, 3) = "2"
            ArrayAcordePadrão(Index, 2, 4) = "3"
            ArrayAcordePadrão(Index, 2, 5) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(#5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 2, 5) = "3"
            ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7(#5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1T"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 4, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7(#5)/E"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "1"
            ArrayAcordePadrão(Index, 1, 4) = "2"
            ArrayAcordePadrão(Index, 1, 5) = "3T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ImagemAcorde(Index) = "C7(#5)/Bb"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "1"
            ArrayAcordePadrão(Index, 1, 4) = "2"
            ArrayAcordePadrão(Index, 1, 5) = "3T"
            ArrayAcordePadrão(Index, 2, 3) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(#5)/Bb"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 4) = "T"
            ArrayAcordePadrão(Index, 2, 1) = "2"
            ArrayAcordePadrão(Index, 2, 3) = "3"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(#5)/Bb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 3, 5) = "3"
            ArrayAcordePadrão(Index, 4, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ImagemAcorde(Index) = "C7(#5/9)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 2, 5) = "3"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(#5/9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 1) = "2T"
            ArrayAcordePadrão(Index, 2, 3) = "3"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(#5/9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 5, 5) = "3"
            ArrayAcordePadrão(Index, 5, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(13)"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 5, 3) = "2"
            ArrayAcordePadrão(Index, 5, 5) = "3"
            ArrayAcordePadrão(Index, 5, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ImagemAcorde(Index) = "C7(13)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "1T"
            ArrayAcordePadrão(Index, 1, 3) = "2"
            ArrayAcordePadrão(Index, 2, 4) = "3"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(13)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ImagemAcorde(Index) = "C7(13)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(9)"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 3, 5) = "3"
            ArrayAcordePadrão(Index, 3, 6) = "3"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ImagemAcorde(Index) = "C7(9)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1"
            ArrayAcordePadrão(Index, 3, 4) = "2"
            ArrayAcordePadrão(Index, 4, 1) = "3T"
            ArrayAcordePadrão(Index, 4, 3) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "1"
            ArrayAcordePadrão(Index, 1, 4) = "2"
            ArrayAcordePadrão(Index, 2, 1) = "3T"
            ArrayAcordePadrão(Index, 2, 3) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ImagemAcorde(Index) = "C7(9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 6) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ImagemAcorde(Index) = "C7(9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 2, 3) = "2T"
            ArrayAcordePadrão(Index, 2, 6) = "3"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste9"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 3, 5) = "3"
            ArrayAcordePadrão(Index, 5, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(9/13)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 4) = "2"
            ArrayAcordePadrão(Index, 4, 1) = "3T"
            ArrayAcordePadrão(Index, 4, 3) = "4"
            ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(9/13)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 5) = "3"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(9/13)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "1"
            ArrayAcordePadrão(Index, 2, 5) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3T"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(b9)"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ImagemAcorde(Index) = "C7(b9)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 4, 1) = "3T"
            ArrayAcordePadrão(Index, 4, 3) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(b9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 2, 2) = "2"
            ArrayAcordePadrão(Index, 3, 1) = "3T"
            ArrayAcordePadrão(Index, 3, 3) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(b9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste6"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 2, 6) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(b9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 2, 6) = "3"
            ArrayAcordePadrão(Index, 3, 2) = "4"
            ImagemAcorde(Index) = "C7(b9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 1, 6) = "2"
            ArrayAcordePadrão(Index, 2, 3) = "3T"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(b9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste9"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ImagemAcorde(Index) = "C7(b5/b9)"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 4, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(#5/b9)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 2, 5) = "3"
            ArrayAcordePadrão(Index, 2, 6) = "4"
            ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(#5/b9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 5, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ImagemAcorde(Index) = "C7(b9/13)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 4, 1) = "3T"
            ArrayAcordePadrão(Index, 4, 3) = "4"
            ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(b9/13)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 2, 6) = "3"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(b9/13)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(#9)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "1"
            ArrayAcordePadrão(Index, 2, 1) = "2T"
            ArrayAcordePadrão(Index, 2, 3) = "3"
            ArrayAcordePadrão(Index, 2, 4) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(#9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 4, 6) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(#9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 4, 6) = "4"
            ArrayAcordePadrão(Index, 7, 3) = ""
            ImagemAcorde(Index) = "C7(#9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 2, 3) = "2T"
            ArrayAcordePadrão(Index, 3, 5) = "3"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(#9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste9"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "1"
            ArrayAcordePadrão(Index, 3, 1) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(#9)/G"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 4, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(#5/#9)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "1"
            ArrayAcordePadrão(Index, 2, 1) = "2T"
            ArrayAcordePadrão(Index, 2, 3) = "3"
            ArrayAcordePadrão(Index, 2, 4) = "3"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(#5/#9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 2, 5) = "3"
            ArrayAcordePadrão(Index, 4, 6) = "4"
            ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(#5/#9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(#9/#11)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 1) = "2T"
            ArrayAcordePadrão(Index, 2, 3) = "3"
            ArrayAcordePadrão(Index, 2, 4) = "4"
            ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(#9/#11)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 3, 3) = "3"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(4)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 3) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(4)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste3"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(4)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 2, 5) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 4, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(4)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "1T"
            ArrayAcordePadrão(Index, 3, 3) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(4/9)"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 7, 1) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(4/9)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 4, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(4/9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste3"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 1) = "3T"
            ArrayAcordePadrão(Index, 3, 3) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(4/9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste6"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 2, 1) = "2T"
            ArrayAcordePadrão(Index, 2, 2) = "3"
            ArrayAcordePadrão(Index, 2, 3) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(4/9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 5) = "2"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(4/9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 3) = "2"
            ArrayAcordePadrão(Index, 3, 6) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ImagemAcorde(Index) = "C7(4/13)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste3"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 2) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(4/13)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 5, 6) = "3"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(4/9/13)"

            'neste precisa desenhar o arco
            Index += 1
            ArrayAcordePadrão(Index, 1, 6) = "1"
            ArrayAcordePadrão(Index, 2, 5) = "1"
            ArrayAcordePadrão(Index, 3, 4) = "2"
            ArrayAcordePadrão(Index, 4, 1) = "3T"
            ArrayAcordePadrão(Index, 4, 3) = "4"
            ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7(4/9/13)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 2, 5) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 3, 3) = "3"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(4/b9)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 3, 1) = "2T"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 3, 3) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(4/b9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste6"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 1) = "3T"
            ArrayAcordePadrão(Index, 3, 3) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "C7(4/b9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste6"

            Index += 1
            ArrayAcordePadrão(Index, 1, 6) = "1"
            ArrayAcordePadrão(Index, 2, 3) = "2T"
            ArrayAcordePadrão(Index, 2, 4) = "3"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7(4/b9)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste9"

            Index += 1
            ArrayAcordePadrão(Index, 2, 4) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 4, 3) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cº"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 2, 1) = "2T"
            ArrayAcordePadrão(Index, 2, 4) = "3"
            ArrayAcordePadrão(Index, 3, 2) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cº"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 1, 5) = "2"
            ArrayAcordePadrão(Index, 2, 1) = "3T"
            ArrayAcordePadrão(Index, 2, 4) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cº"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 1, 5) = "2"
            ArrayAcordePadrão(Index, 2, 4) = "3"
            ArrayAcordePadrão(Index, 2, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cº"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 2, 4) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 4, 5) = "3"
            ArrayAcordePadrão(Index, 4, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ArrayAcordePadrão(Index, 8, 1) = "*"
            ImagemAcorde(Index) = "Cº(b13)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 2, 1) = "2T"
            ArrayAcordePadrão(Index, 2, 4) = "3"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cº(b13)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "1T"
            ArrayAcordePadrão(Index, 4, 3) = "2"
            ArrayAcordePadrão(Index, 4, 4) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cº(7M)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 1) = "2T"
            ArrayAcordePadrão(Index, 2, 4) = "3"
            ArrayAcordePadrão(Index, 3, 2) = "4"
            ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "Cº(7M)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

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

            Index += 1 : AjustePosição(Index) = 3
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

            Index += 1 : AjustePosição(Index) = 5
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

            Index += 1 : AjustePosição(Index) = 8
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

            Index += 1 : AjustePosição(Index) = 10
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

            Index += 1 : AjustePosição(Index) = 3
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

            Index += 1 : AjustePosição(Index) = 5
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

            Index += 1 : AjustePosição(Index) = 8
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

            Index += 1 : AjustePosição(Index) = 10
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

            Index += 1 : AjustePosição(Index) = 3
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

            Index += 1 : AjustePosição(Index) = 5
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

            Index += 1 : AjustePosição(Index) = 8
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

            Index += 1 : AjustePosição(Index) = 10
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

            Index += 1 : AjustePosição(Index) = 3
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

            Index += 1 : AjustePosição(Index) = 5
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

            Index += 1 : AjustePosição(Index) = 8
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

            Index += 1 : AjustePosição(Index) = 10
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
            ArrayAcordePadrão(Index, 1, 5) = "1T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ImagemAcorde(Index) = "C/G"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1T"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C/E"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 3, 1) = "3"
            ArrayAcordePadrão(Index, 3, 2) = "4T"
            ArrayAcordePadrão(Index, 7, 4) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C/G"

            Index += 1
            ArrayAcordePadrão(Index, 3, 6) = "1"
            ArrayAcordePadrão(Index, 5, 4) = "3T"
            ArrayAcordePadrão(Index, 5, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ImagemAcorde(Index) = "C"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 5 : ArrayAcordePadrão(Index, 1, 4) = "T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 1, 4) = "T"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 4, 1) = "4T"
            ArrayAcordePadrão(Index, 7, 4) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 6) = "T"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ImagemAcorde(Index) = "C/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 3) = "3T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "2"
            ArrayAcordePadrão(Index, 2, 2) = "3"
            ArrayAcordePadrão(Index, 2, 3) = "4T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste9"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 3 : ArrayAcordePadrão(Index, 1, 3) = "T"
            ArrayAcordePadrão(Index, 3, 1) = "3"
            ArrayAcordePadrão(Index, 7, 4) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 1, 6) = "2"
            ArrayAcordePadrão(Index, 2, 5) = "3T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ImagemAcorde(Index) = "C/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste12"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 2, 5) = "2T"
            ArrayAcordePadrão(Index, 3, 3) = "3"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste12"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 3, 3) = "3"
            ArrayAcordePadrão(Index, 4, 2) = "4T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C"
            ArrayAcordePadrão(Index, 1, 0) = "Traste12"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 1, 5) = "2T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm/Eb"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "3T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 3, 1) = "3"
            ArrayAcordePadrão(Index, 3, 2) = "4T"
            ArrayAcordePadrão(Index, 7, 4) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm/G"

            Index += 1
            ArrayAcordePadrão(Index, 3, 6) = "1"
            ArrayAcordePadrão(Index, 4, 5) = "2"
            ArrayAcordePadrão(Index, 5, 4) = "3T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ImagemAcorde(Index) = "Cm"

            Index += 1
            ArrayAcordePadrão(Index, 4, 5) = "1"
            ArrayAcordePadrão(Index, 5, 3) = "2"
            ArrayAcordePadrão(Index, 5, 4) = "3T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm/G"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 1, 4) = "T"
            ArrayAcordePadrão(Index, 2, 2) = "2"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm/Eb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 2, 2) = "2"
            ArrayAcordePadrão(Index, 4, 1) = "4T"
            ArrayAcordePadrão(Index, 7, 4) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 6) = "T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ImagemAcorde(Index) = "Cm/Eb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "P" : LinhaFinalDaSeta(Index) = 5
            ArrayAcordePadrão(Index, 3, 3) = "3T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "4"
            ArrayAcordePadrão(Index, 3, 3) = "3T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 3 : ArrayAcordePadrão(Index, 1, 3) = "T"
            ArrayAcordePadrão(Index, 2, 1) = "2"
            ArrayAcordePadrão(Index, 7, 4) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm/Eb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 6) = "1"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 5) = "3T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ImagemAcorde(Index) = "Cm/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste11"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 2, 5) = "3T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm/Eb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste12"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 4, 2) = "4T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm"
            ArrayAcordePadrão(Index, 1, 0) = "Traste12"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ImagemAcorde(Index) = "C4/G"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1T"
            ArrayAcordePadrão(Index, 3, 3) = "3"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C4/F"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "1T"
            ArrayAcordePadrão(Index, 3, 3) = "2"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C4"

            Index += 1
            ArrayAcordePadrão(Index, 3, 1) = "P" : LinhaFinalDaSeta(Index) = 3 : ArrayAcordePadrão(Index, 3, 2) = "T"
            ArrayAcordePadrão(Index, 7, 4) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C4/G"

            Index += 1
            ArrayAcordePadrão(Index, 3, 6) = "1"
            ArrayAcordePadrão(Index, 5, 4) = "3T"
            ArrayAcordePadrão(Index, 6, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ImagemAcorde(Index) = "C4"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 1, 4) = "T"
            ArrayAcordePadrão(Index, 2, 5) = "2"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C4/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 1, 4) = "T"
            ArrayAcordePadrão(Index, 4, 2) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C4/F"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 4, 1) = "3T"
            ArrayAcordePadrão(Index, 4, 2) = "4"
            ArrayAcordePadrão(Index, 7, 4) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C4"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 6) = "T"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ImagemAcorde(Index) = "C4/F"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1"
            ArrayAcordePadrão(Index, 3, 3) = "3T"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C4"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 1, 3) = "T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C4/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 3 : ArrayAcordePadrão(Index, 1, 3) = "T"
            ArrayAcordePadrão(Index, 4, 1) = "4"
            ArrayAcordePadrão(Index, 7, 4) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C4/F"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 2, 5) = "2T"
            ArrayAcordePadrão(Index, 2, 6) = "3"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
            ImagemAcorde(Index) = "C4/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste12"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 2, 5) = "2T"
            ArrayAcordePadrão(Index, 4, 3) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C4/F"
            ArrayAcordePadrão(Index, 1, 0) = "Traste12"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 4, 2) = "3T"
            ArrayAcordePadrão(Index, 4, 3) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C4"
            ArrayAcordePadrão(Index, 1, 0) = "Traste12"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1T"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C/E"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1T"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C/E"

            Index += 1
            ArrayAcordePadrão(Index, 3, 6) = "1"
            ArrayAcordePadrão(Index, 5, 3) = "2"
            ArrayAcordePadrão(Index, 5, 4) = "3T"
            ArrayAcordePadrão(Index, 5, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C/G"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 4, 6) = "4T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 3) = "3T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 3, 4) = "2"
            ArrayAcordePadrão(Index, 3, 6) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 5
            ArrayAcordePadrão(Index, 5, 3) = "2"
            ArrayAcordePadrão(Index, 5, 4) = "3"
            ArrayAcordePadrão(Index, 5, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 5 : ArrayAcordePadrão(Index, 1, 4) = "T"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 5 : ArrayAcordePadrão(Index, 1, 4) = "T"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 3, 3) = "4T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "1"
            ArrayAcordePadrão(Index, 3, 4) = "2"
            ArrayAcordePadrão(Index, 4, 5) = "3T"
            ArrayAcordePadrão(Index, 5, 3) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 3, 1) = "3"
            ArrayAcordePadrão(Index, 3, 2) = "4T"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C/G"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "1"
            ArrayAcordePadrão(Index, 3, 3) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3T"
            ArrayAcordePadrão(Index, 5, 2) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste3"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 4
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 4, 1) = "4T"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "1T"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 3, 3) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 3 : ArrayAcordePadrão(Index, 1, 3) = "T"
            ArrayAcordePadrão(Index, 3, 1) = "3"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "P" : LinhaFinalDaSeta(Index) = 4
            ArrayAcordePadrão(Index, 3, 3) = "3"
            ArrayAcordePadrão(Index, 4, 2) = "4T"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste12"

            Index += 1
            ArrayAcordePadrão(Index, 1, 6) = "1"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 3) = "3"
            ArrayAcordePadrão(Index, 3, 5) = "4T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "Cm/Eb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste11"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 2, 5) = "3T"
            ArrayAcordePadrão(Index, 4, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "Cm/Eb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste12"

            Index += 1
            ArrayAcordePadrão(Index, 3, 6) = "1"
            ArrayAcordePadrão(Index, 4, 5) = "2"
            ArrayAcordePadrão(Index, 5, 3) = "3"
            ArrayAcordePadrão(Index, 5, 4) = "4T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "Cm/G"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 4, 4) = "2"
            ArrayAcordePadrão(Index, 4, 5) = "3"
            ArrayAcordePadrão(Index, 4, 6) = "4T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "Cm/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 3) = "3T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "Cm"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 2, 6) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "Cm"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 1, 5) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "4T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "1T"
            ArrayAcordePadrão(Index, 4, 5) = "2"
            ArrayAcordePadrão(Index, 5, 3) = "3"
            ArrayAcordePadrão(Index, 5, 4) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm"

            Index += 1
            ArrayAcordePadrão(Index, 4, 5) = "1"
            ArrayAcordePadrão(Index, 5, 3) = "2"
            ArrayAcordePadrão(Index, 5, 4) = "3T"
            ArrayAcordePadrão(Index, 6, 2) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm/Eb"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 1, 4) = "T"
            ArrayAcordePadrão(Index, 2, 2) = "2"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm/Eb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "P" : LinhaFinalDaSeta(Index) = 5
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 3, 3) = "4T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "1"
            ArrayAcordePadrão(Index, 3, 4) = "2"
            ArrayAcordePadrão(Index, 4, 3) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 3, 1) = "3"
            ArrayAcordePadrão(Index, 3, 2) = "4T"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm/G"

            Index += 1
            ArrayAcordePadrão(Index, 3, 1) = "1"
            ArrayAcordePadrão(Index, 5, 3) = "2"
            ArrayAcordePadrão(Index, 5, 4) = "3T"
            ArrayAcordePadrão(Index, 6, 2) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm/G"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 4
            ArrayAcordePadrão(Index, 2, 2) = "2"
            ArrayAcordePadrão(Index, 4, 1) = "4T"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 4
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 3, 3) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 3 : ArrayAcordePadrão(Index, 1, 3) = "T"
            ArrayAcordePadrão(Index, 2, 1) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm/Eb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "1"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 3) = "3"
            ArrayAcordePadrão(Index, 5, 2) = "4T"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm/Eb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste11"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 1, 5) = "2T"
            ArrayAcordePadrão(Index, 2, 3) = "3"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C(#5)/E"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 5) = "T"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 4, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C(#5)/E"

            Index += 1
            ArrayAcordePadrão(Index, 4, 6) = "1"
            ArrayAcordePadrão(Index, 5, 4) = "2T"
            ArrayAcordePadrão(Index, 5, 5) = "3"
            ArrayAcordePadrão(Index, 6, 3) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C(#5)/G#"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 4, 6) = "4T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C(#5)/G#"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 6) = "1"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 2, 5) = "3"
            ArrayAcordePadrão(Index, 3, 3) = "4T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C(#5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 3, 6) = "2"
            ArrayAcordePadrão(Index, 4, 4) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C(#5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "P" : LinhaFinalDaSeta(Index) = 5
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C(#5)"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "1T"
            ArrayAcordePadrão(Index, 5, 4) = "2"
            ArrayAcordePadrão(Index, 5, 5) = "3"
            ArrayAcordePadrão(Index, 6, 3) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C(#5)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 5 : ArrayAcordePadrão(Index, 1, 4) = "T"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C(#5)/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1T"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 5, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C(#5)/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "P" : LinhaFinalDaSeta(Index) = 5
            ArrayAcordePadrão(Index, 2, 3) = "2T"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C(#5)/G#"
            ArrayAcordePadrão(Index, 1, 0) = "Traste9"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "1"
            ArrayAcordePadrão(Index, 3, 4) = "2"
            ArrayAcordePadrão(Index, 3, 5) = "3T"
            ArrayAcordePadrão(Index, 4, 3) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C(#5)/G#"
            ArrayAcordePadrão(Index, 1, 0) = "Traste11"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3T"
            ArrayAcordePadrão(Index, 4, 1) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C(#5)/G#"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "1"
            ArrayAcordePadrão(Index, 2, 4) = "2T"
            ArrayAcordePadrão(Index, 3, 3) = "3"
            ArrayAcordePadrão(Index, 4, 2) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C(#5)/G#"
            ArrayAcordePadrão(Index, 1, 0) = "Traste4"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 4, 1) = "4T"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C(#5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "1T"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 3) = "3"
            ArrayAcordePadrão(Index, 4, 2) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C(#5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 2, 2) = "2"
            ArrayAcordePadrão(Index, 3, 1) = "3"
            ArrayAcordePadrão(Index, 4, 4) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C(#5)/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "1"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 3) = "3"
            ArrayAcordePadrão(Index, 4, 2) = "4T"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C(#5)/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste12"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 3) = "3"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C4/F"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1T"
            ArrayAcordePadrão(Index, 3, 3) = "3"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C4/F"

            Index += 1
            ArrayAcordePadrão(Index, 3, 6) = "1"
            ArrayAcordePadrão(Index, 5, 3) = "2"
            ArrayAcordePadrão(Index, 5, 4) = "3T"
            ArrayAcordePadrão(Index, 6, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C4/G"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 5) = "2"
            ArrayAcordePadrão(Index, 4, 6) = "4T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C4/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 3) = "3T"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C4"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 3, 4) = "2"
            ArrayAcordePadrão(Index, 4, 5) = "3"
            ArrayAcordePadrão(Index, 4, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C4"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "3T"
            ArrayAcordePadrão(Index, 3, 3) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C4"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "1T"
            ArrayAcordePadrão(Index, 5, 3) = "2"
            ArrayAcordePadrão(Index, 5, 4) = "3"
            ArrayAcordePadrão(Index, 6, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C4"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 5 : ArrayAcordePadrão(Index, 1, 4) = "T"
            ArrayAcordePadrão(Index, 2, 5) = "2"
            ArrayAcordePadrão(Index, 4, 2) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C4/F"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 1, 4) = "T"
            ArrayAcordePadrão(Index, 4, 2) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C4/F"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2"
            ArrayAcordePadrão(Index, 3, 3) = "3T"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C4/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "1"
            ArrayAcordePadrão(Index, 3, 4) = "2"
            ArrayAcordePadrão(Index, 4, 5) = "3T"
            ArrayAcordePadrão(Index, 6, 3) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C4/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 3, 1) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3T"
            ArrayAcordePadrão(Index, 3, 3) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C4/G"

            Index += 1
            ArrayAcordePadrão(Index, 3, 1) = "P" : LinhaFinalDaSeta(Index) = 4
            ArrayAcordePadrão(Index, 5, 4) = "3T"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C4/G"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 4
            ArrayAcordePadrão(Index, 4, 1) = "3T"
            ArrayAcordePadrão(Index, 4, 2) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C4"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "1T"
            ArrayAcordePadrão(Index, 3, 2) = "2"
            ArrayAcordePadrão(Index, 3, 3) = "3"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C4"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 3 : ArrayAcordePadrão(Index, 1, 3) = "T"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 4, 1) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C4/F"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 2, 1) = "2"
            ArrayAcordePadrão(Index, 4, 2) = "3T"
            ArrayAcordePadrão(Index, 4, 3) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C4/F"
            ArrayAcordePadrão(Index, 1, 0) = "Traste12"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1T"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 2, 4) = "3"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C6/E"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 4) = "T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C6/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 2, 5) = "2"
            ArrayAcordePadrão(Index, 2, 6) = "3T"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C6/A = Am7"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C6"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1T"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C6/A = Am7"

            Index += 1
            ArrayAcordePadrão(Index, 2, 4) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 5, 3) = "3"
            ArrayAcordePadrão(Index, 5, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C6"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1T"
            ArrayAcordePadrão(Index, 3, 2) = "2"
            ArrayAcordePadrão(Index, 3, 3) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C6/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 2, 2) = "2"
            ArrayAcordePadrão(Index, 2, 3) = "3T"
            ArrayAcordePadrão(Index, 2, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C6/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste9"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "P" : LinhaFinalDaSeta(Index) = 4
            ArrayAcordePadrão(Index, 3, 1) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3T"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C6/G"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "P" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 1, 4) = "T"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C6/A = Am7"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 2, 1) = "2T"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 4, 2) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C6"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 3, 1) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C6/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 5) = "T"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "Cm6/Eb"

            Index += 1
            ArrayAcordePadrão(Index, 4, 5) = "1"
            ArrayAcordePadrão(Index, 5, 3) = "2"
            ArrayAcordePadrão(Index, 5, 4) = "3T"
            ArrayAcordePadrão(Index, 5, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "Cm6/G"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 2, 5) = "3"
            ArrayAcordePadrão(Index, 2, 6) = "4T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "Cm6/A = Am7(b5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "PT" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 6) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "Cm6"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 1, 5) = "2T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm6/A = Am7(b5)"

            Index += 1
            ArrayAcordePadrão(Index, 2, 4) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 4, 5) = "3"
            ArrayAcordePadrão(Index, 5, 3) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm6"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1T"
            ArrayAcordePadrão(Index, 2, 2) = "2"
            ArrayAcordePadrão(Index, 3, 3) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm6/Eb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2"
            ArrayAcordePadrão(Index, 3, 3) = "3T"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm6/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 1) = "3"
            ArrayAcordePadrão(Index, 3, 2) = "4T"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm6/G"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "P" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 1, 4) = "T"
            ArrayAcordePadrão(Index, 2, 2) = "2"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm6/A = Am7(b5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 2, 1) = "2T"
            ArrayAcordePadrão(Index, 2, 4) = "3"
            ArrayAcordePadrão(Index, 4, 2) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm6"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 2, 1) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Cm6/Eb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1T"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7/E"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 4) = "T"
            ArrayAcordePadrão(Index, 2, 6) = "2"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 6) = "T"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7/Bb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 2, 5) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "1"
            ArrayAcordePadrão(Index, 1, 5) = "2T"
            ArrayAcordePadrão(Index, 2, 3) = "3"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7/Bb"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 5
            ArrayAcordePadrão(Index, 5, 3) = "3"
            ArrayAcordePadrão(Index, 5, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1T"
            ArrayAcordePadrão(Index, 3, 2) = "2"
            ArrayAcordePadrão(Index, 4, 3) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 2, 2) = "2"
            ArrayAcordePadrão(Index, 2, 3) = "3T"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste9"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "1"
            ArrayAcordePadrão(Index, 3, 1) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3T"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7/G"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 1, 4) = "T"
            ArrayAcordePadrão(Index, 2, 1) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7/Bb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 4
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 3, 1) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 4, 2) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1T"
            ArrayAcordePadrão(Index, 3, 3) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7(4)/F"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 4) = "T"
            ArrayAcordePadrão(Index, 2, 5) = "2"
            ArrayAcordePadrão(Index, 2, 6) = "3"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7(4)/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 6) = "T"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7(4)/Bb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 2, 5) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 4, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7(4)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "1"
            ArrayAcordePadrão(Index, 1, 5) = "2T"
            ArrayAcordePadrão(Index, 3, 3) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(4)/Bb"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 5
            ArrayAcordePadrão(Index, 5, 3) = "3"
            ArrayAcordePadrão(Index, 6, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(4)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1T"
            ArrayAcordePadrão(Index, 4, 2) = "2"
            ArrayAcordePadrão(Index, 4, 3) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(4)/F"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 5 : ArrayAcordePadrão(Index, 1, 3) = "T"
            ArrayAcordePadrão(Index, 2, 5) = "2"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(4)/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 3, 1) = "P" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 3, 2) = "T"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(4)/G"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 1, 4) = "T"
            ArrayAcordePadrão(Index, 2, 1) = "2"
            ArrayAcordePadrão(Index, 4, 2) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(4)/Bb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 4
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(4)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 3, 4) = "2"
            ArrayAcordePadrão(Index, 4, 1) = "3"
            ArrayAcordePadrão(Index, 4, 2) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(4)/F"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1T"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 2, 6) = "3"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7(b5)/E"

            Index += 1
            ArrayAcordePadrão(Index, 4, 3) = "1"
            ArrayAcordePadrão(Index, 5, 4) = "2T"
            ArrayAcordePadrão(Index, 5, 5) = "3"
            ArrayAcordePadrão(Index, 6, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7(b5)/Gb"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 2, 6) = "3T"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7(b5)/Bb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 2, 5) = "3"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7(b5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2"
            ArrayAcordePadrão(Index, 3, 5) = "3T"
            ArrayAcordePadrão(Index, 4, 3) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(b5)/Bb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste11"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 5
            ArrayAcordePadrão(Index, 4, 3) = "2"
            ArrayAcordePadrão(Index, 5, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(b5)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1T"
            ArrayAcordePadrão(Index, 3, 2) = "2"
            ArrayAcordePadrão(Index, 3, 5) = "3"
            ArrayAcordePadrão(Index, 4, 3) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(b5)/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "1"
            ArrayAcordePadrão(Index, 1, 4) = "2"
            ArrayAcordePadrão(Index, 2, 3) = "3T"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(b5)/Gb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste9"

            Index += 1
            ArrayAcordePadrão(Index, 2, 1) = "1"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3T"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(b5)/Gb"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 2, 4) = "2T"
            ArrayAcordePadrão(Index, 3, 1) = "3"
            ArrayAcordePadrão(Index, 4, 2) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(b5)/Bb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste4"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "1T"
            ArrayAcordePadrão(Index, 1, 3) = "2"
            ArrayAcordePadrão(Index, 2, 2) = "3"
            ArrayAcordePadrão(Index, 2, 4) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(b5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 1) = "3"
            ArrayAcordePadrão(Index, 4, 2) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(b5)/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1T"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 4, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7(#5)/E"

            Index += 1
            ArrayAcordePadrão(Index, 5, 4) = "1T"
            ArrayAcordePadrão(Index, 5, 5) = "2"
            ArrayAcordePadrão(Index, 6, 3) = "3"
            ArrayAcordePadrão(Index, 6, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7(#5)/G#"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 1, 6) = "2T"
            ArrayAcordePadrão(Index, 2, 4) = "3"
            ArrayAcordePadrão(Index, 2, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7(#5)/Bb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 2, 5) = "2"
            ArrayAcordePadrão(Index, 3, 6) = "3"
            ArrayAcordePadrão(Index, 4, 4) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7(#5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 5 : ArrayAcordePadrão(Index, 1, 5) = "T"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(#5)/Bb"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 5
            ArrayAcordePadrão(Index, 5, 5) = "3"
            ArrayAcordePadrão(Index, 6, 3) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(#5)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1T"
            ArrayAcordePadrão(Index, 3, 2) = "2"
            ArrayAcordePadrão(Index, 4, 3) = "3"
            ArrayAcordePadrão(Index, 5, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(#5)/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 2, 3) = "2T"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(#5)/G#"
            ArrayAcordePadrão(Index, 1, 0) = "Traste9"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 4, 1) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(#5)/G#"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1T"
            ArrayAcordePadrão(Index, 2, 1) = "2"
            ArrayAcordePadrão(Index, 2, 3) = "3"
            ArrayAcordePadrão(Index, 3, 2) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(#5)/Bb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 4
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 4, 2) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(#5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 3, 1) = "2"
            ArrayAcordePadrão(Index, 4, 2) = "3"
            ArrayAcordePadrão(Index, 4, 4) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7(#5)/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1T"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 3, 6) = "3"
            ArrayAcordePadrão(Index, 4, 4) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7M/E"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 4) = "T"
            ArrayAcordePadrão(Index, 3, 6) = "3"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7M/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 6) = "T"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 2, 4) = "3"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7M/B"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 3, 4) = "2"
            ArrayAcordePadrão(Index, 3, 5) = "3"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7M"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1T"
            ArrayAcordePadrão(Index, 2, 2) = "2"
            ArrayAcordePadrão(Index, 2, 3) = "3"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M/B"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "1T"
            ArrayAcordePadrão(Index, 4, 4) = "2"
            ArrayAcordePadrão(Index, 5, 3) = "3"
            ArrayAcordePadrão(Index, 5, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1T"
            ArrayAcordePadrão(Index, 3, 2) = "2"
            ArrayAcordePadrão(Index, 4, 5) = "3"
            ArrayAcordePadrão(Index, 5, 3) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 2, 2) = "2"
            ArrayAcordePadrão(Index, 2, 3) = "3T"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste9"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "1"
            ArrayAcordePadrão(Index, 3, 1) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3T"
            ArrayAcordePadrão(Index, 4, 4) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M/G"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 1, 4) = "T"
            ArrayAcordePadrão(Index, 3, 1) = "3"
            ArrayAcordePadrão(Index, 3, 2) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M/B"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "1T"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 2, 4) = "3"
            ArrayAcordePadrão(Index, 3, 2) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 3, 1) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 5, 2) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1T"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 2, 6) = "3"
            ArrayAcordePadrão(Index, 4, 4) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7M(b5)/E"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 2, 4) = "2T"
            ArrayAcordePadrão(Index, 2, 5) = "3"
            ArrayAcordePadrão(Index, 4, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7M(b5)/Gb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste4"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1"
            ArrayAcordePadrão(Index, 2, 6) = "2T"
            ArrayAcordePadrão(Index, 3, 3) = "3"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7M(b5)/B"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 5) = "3"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7M(b5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 3, 5) = "2T"
            ArrayAcordePadrão(Index, 4, 2) = "3"
            ArrayAcordePadrão(Index, 4, 3) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M(b5)/B"
            ArrayAcordePadrão(Index, 1, 0) = "Traste11"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "1T"
            ArrayAcordePadrão(Index, 4, 3) = "2"
            ArrayAcordePadrão(Index, 4, 4) = "3"
            ArrayAcordePadrão(Index, 5, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M(b5)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1T"
            ArrayAcordePadrão(Index, 3, 2) = "2"
            ArrayAcordePadrão(Index, 3, 5) = "3"
            ArrayAcordePadrão(Index, 5, 3) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M(b5)/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 4
            ArrayAcordePadrão(Index, 2, 3) = "2T"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M(b5)/Gb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste9"

            Index += 1
            ArrayAcordePadrão(Index, 2, 1) = "P" : LinhaFinalDaSeta(Index) = 4
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 4, 4) = "3"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M(b5)/Gb"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 2, 4) = "2T"
            ArrayAcordePadrão(Index, 4, 1) = "3"
            ArrayAcordePadrão(Index, 4, 2) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M(b5)/B"
            ArrayAcordePadrão(Index, 1, 0) = "Traste4"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "1T"
            ArrayAcordePadrão(Index, 2, 2) = "2"
            ArrayAcordePadrão(Index, 2, 3) = "3"
            ArrayAcordePadrão(Index, 2, 4) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M(b5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 3, 1) = "3"
            ArrayAcordePadrão(Index, 5, 2) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M(b5)/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 5) = "1T"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 4, 4) = "3"
            ArrayAcordePadrão(Index, 4, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7M(#5)/E"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "PT" : LinhaFinalDaSeta(Index) = 5
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 3, 6) = "3"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7M(#5)/G#"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 6) = "1T"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 2, 4) = "3"
            ArrayAcordePadrão(Index, 2, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7M(#5)/B"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 3, 5) = "2"
            ArrayAcordePadrão(Index, 3, 6) = "3"
            ArrayAcordePadrão(Index, 4, 4) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "C7M(#5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "P" : LinhaFinalDaSeta(Index) = 5 : ArrayAcordePadrão(Index, 1, 5) = "T"
            ArrayAcordePadrão(Index, 2, 2) = "2"
            ArrayAcordePadrão(Index, 2, 3) = "3"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M(#5)/B"

            Index += 1
            ArrayAcordePadrão(Index, 3, 2) = "1T"
            ArrayAcordePadrão(Index, 4, 4) = "2"
            ArrayAcordePadrão(Index, 5, 5) = "3"
            ArrayAcordePadrão(Index, 6, 3) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M(#5)"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1T"
            ArrayAcordePadrão(Index, 3, 2) = "2"
            ArrayAcordePadrão(Index, 5, 3) = "3"
            ArrayAcordePadrão(Index, 5, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M(#5)/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 2, 3) = "2T"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M(#5)/G#"
            ArrayAcordePadrão(Index, 1, 0) = "Traste9"

            Index += 1
            ArrayAcordePadrão(Index, 2, 3) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2T"
            ArrayAcordePadrão(Index, 4, 1) = "3"
            ArrayAcordePadrão(Index, 4, 4) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M(#5)/G#"

            Index += 1
            ArrayAcordePadrão(Index, 1, 4) = "1T"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 3, 1) = "3"
            ArrayAcordePadrão(Index, 3, 2) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M(#5)/B"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1
            ArrayAcordePadrão(Index, 1, 1) = "1T"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 2, 4) = "3"
            ArrayAcordePadrão(Index, 4, 2) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M(#5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 3, 1) = "2"
            ArrayAcordePadrão(Index, 4, 4) = "3"
            ArrayAcordePadrão(Index, 5, 2) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "C7M(#5)/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 1, 5) = "1"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 2, 4) = "3T"
            ArrayAcordePadrão(Index, 4, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "Am(7M)/E"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 6) = "T"
            ArrayAcordePadrão(Index, 2, 3) = "3"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "Am(7M)/G#"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 2, 6) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 3, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "Am(7M)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 5) = "T"
            ArrayAcordePadrão(Index, 3, 6) = "3"
            ArrayAcordePadrão(Index, 4, 4) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "Am(7M)/C"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 0, 2) = "T"
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 1, 5) = "2"
            ArrayAcordePadrão(Index, 2, 3) = "3"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Am(7M)"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 2, 4) = "1T"
            ArrayAcordePadrão(Index, 3, 2) = "2"
            ArrayAcordePadrão(Index, 5, 5) = "3"
            ArrayAcordePadrão(Index, 6, 3) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Am(7M)/C"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2"
            ArrayAcordePadrão(Index, 3, 3) = "3T"
            ArrayAcordePadrão(Index, 5, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Am(7M)/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 2, 5) = "3T"
            ArrayAcordePadrão(Index, 3, 2) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Am(7M)/G#"
            ArrayAcordePadrão(Index, 1, 0) = "Traste9"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 2, 3) = "P" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 2, 4) = "T"
            ArrayAcordePadrão(Index, 3, 2) = "2"
            ArrayAcordePadrão(Index, 4, 1) = "3"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Am(7M)/G#"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 4
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Am(7M)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 2, 1) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 5, 2) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Am(7M)/C"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 3, 1) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3T"
            ArrayAcordePadrão(Index, 4, 4) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Am(7M)/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 1, 5) = "1"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 2, 4) = "3T"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "Am7/E"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 6) = "T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "Am7/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 2, 5) = "2"
            ArrayAcordePadrão(Index, 2, 6) = "3"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "Am7"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 5) = "T"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "Am7/C"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 0, 2) = "T"
            ArrayAcordePadrão(Index, 1, 5) = "1"
            ArrayAcordePadrão(Index, 2, 3) = "2"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Am7"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 2, 4) = "1T"
            ArrayAcordePadrão(Index, 3, 2) = "2"
            ArrayAcordePadrão(Index, 5, 3) = "3"
            ArrayAcordePadrão(Index, 5, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Am7/C"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2"
            ArrayAcordePadrão(Index, 3, 3) = "3T"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Am7/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 2, 2) = "2"
            ArrayAcordePadrão(Index, 2, 3) = "3"
            ArrayAcordePadrão(Index, 2, 5) = "4T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Am7/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste9"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 2, 3) = "P" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 2, 4) = "T"
            ArrayAcordePadrão(Index, 3, 1) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Am7/G"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 4
            ArrayAcordePadrão(Index, 3, 2) = "3"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Am7"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 2, 1) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 4, 2) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Am7/C"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 3, 1) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3T"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Am7/E"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6
            ArrayAcordePadrão(Index, 2, 4) = "2T"
            ArrayAcordePadrão(Index, 3, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "Am7(b5)/Eb"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 4, 5) = "1"
            ArrayAcordePadrão(Index, 5, 3) = "2"
            ArrayAcordePadrão(Index, 5, 4) = "3"
            ArrayAcordePadrão(Index, 5, 6) = "4T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "Am7(b5)/G"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 2, 4) = "2"
            ArrayAcordePadrão(Index, 2, 5) = "3"
            ArrayAcordePadrão(Index, 2, 6) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "Am7(b5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 5) = "T"
            ArrayAcordePadrão(Index, 2, 6) = "2"
            ArrayAcordePadrão(Index, 3, 4) = "3"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
            ImagemAcorde(Index) = "Am7(b5)/C"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 0, 2) = "T"
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 1, 5) = "2"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Am7(b5)"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 2, 4) = "1T"
            ArrayAcordePadrão(Index, 3, 2) = "2"
            ArrayAcordePadrão(Index, 4, 5) = "3"
            ArrayAcordePadrão(Index, 5, 3) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Am7(b5)/C"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 2, 2) = "2"
            ArrayAcordePadrão(Index, 3, 3) = "3T"
            ArrayAcordePadrão(Index, 4, 5) = "4"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Am7(b5)/Eb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 1, 4) = "1"
            ArrayAcordePadrão(Index, 3, 2) = "2"
            ArrayAcordePadrão(Index, 3, 3) = "3"
            ArrayAcordePadrão(Index, 3, 5) = "4T"
            ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Am7(b5)/G"
            ArrayAcordePadrão(Index, 1, 0) = "Traste8"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 2, 4) = "2T"
            ArrayAcordePadrão(Index, 3, 1) = "3"
            ArrayAcordePadrão(Index, 3, 2) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Am7(b5)/G"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 4
            ArrayAcordePadrão(Index, 2, 2) = "2"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Am7(b5)"
            ArrayAcordePadrão(Index, 1, 0) = "Traste5"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 1, 3) = "1T"
            ArrayAcordePadrão(Index, 2, 1) = "2"
            ArrayAcordePadrão(Index, 2, 4) = "3"
            ArrayAcordePadrão(Index, 4, 2) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Am7(b5)/C"
            ArrayAcordePadrão(Index, 1, 0) = "Traste7"

            Index += 1 : AjustePosição(Index) = 3
            ArrayAcordePadrão(Index, 1, 3) = "1"
            ArrayAcordePadrão(Index, 2, 1) = "2"
            ArrayAcordePadrão(Index, 3, 2) = "3T"
            ArrayAcordePadrão(Index, 3, 4) = "4"
            ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
            ImagemAcorde(Index) = "Am7(b5)/Eb"
            ArrayAcordePadrão(Index, 1, 0) = "Traste10"

            'MsgBox("Qtde Acordes: " & Index)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub SubstituirString()

        Try

            For i = 0 To 27
                For ii = 0 To 6
                    If IntervalosAcorde(i, ii) = StringSubstituição(0) Then IntervalosAcorde(i, ii) = StringSubstituição(1)
                Next
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        My.Settings.NovoValorTipoDeJogoTelaEstudoIntervalos = True
        Desenhar()
    End Sub

    Private Sub ComboBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.Click, ComboBox6.Click, ComboBox5.Click, ComboBox4.Click, ComboBox3.Click, ComboBox2.Click
        Dim pCombo As ComboBox = DirectCast(sender, ComboBox)
        pCombo.DroppedDown = True
    End Sub

    Private Sub ComboBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.TextChanged
        Try
            If ComboBox1.Text = IntervalosDoAcordeGerado(1) Then
                Label1.BackColor = Color.Green
            Else
                Label1.BackColor = Color.Red
            End If
            Label1.ForeColor = Color.White
            VerificarSeExercícioEstáCorreto()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub ComboBox2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.TextChanged
        Try
            If ComboBox2.Text = IntervalosDoAcordeGerado(2) Then
                Label2.BackColor = Color.Green
            Else
                Label2.BackColor = Color.Red
            End If
            Label2.ForeColor = Color.White
            VerificarSeExercícioEstáCorreto()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub ComboBox3_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.TextChanged
        Try
            If ComboBox3.Text = IntervalosDoAcordeGerado(3) Then
                Label3.BackColor = Color.Green
            Else
                Label3.BackColor = Color.Red
            End If
            Label3.ForeColor = Color.White
            VerificarSeExercícioEstáCorreto()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub ComboBox4_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox4.TextChanged
        Try
            If ComboBox4.Text = IntervalosDoAcordeGerado(4) Then
                Label4.BackColor = Color.Green
            Else
                Label4.BackColor = Color.Red
            End If
            Label4.ForeColor = Color.White
            VerificarSeExercícioEstáCorreto()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub ComboBox5_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox5.TextChanged
        Try
            If ComboBox5.Text = IntervalosDoAcordeGerado(5) Then
                Label5.BackColor = Color.Green
            Else
                Label5.BackColor = Color.Red
            End If
            Label5.ForeColor = Color.White
            VerificarSeExercícioEstáCorreto()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub ComboBox6_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox6.TextChanged
        Try
            If ComboBox6.Text = IntervalosDoAcordeGerado(6) Then
                Label6.BackColor = Color.Green
            Else
                Label6.BackColor = Color.Red
            End If
            Label6.ForeColor = Color.White
            VerificarSeExercícioEstáCorreto()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            If Me.Width = 247 Then
                Me.Width = 490
            Else
                Me.Width = 247
            End If

            My.Settings.NovoValorLarguraTelaEstudoIntervalos = Me.Width
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub EstudoIntervalosAcordes_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        Try
            If Me.Width = 247 Then
                Button2.Text = "Mostrar Intervalos"
            Else
                Button2.Text = "Ocultar Intervalos"
            End If
            Me.Left = CInt((Screen.PrimaryScreen.Bounds.Width - Me.Width) / 2)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub VerificarSeExercícioEstáCorreto()
        Try

            If (Label1.BackColor = Color.Green OrElse Not ComboBox1.Visible) AndAlso (Label2.BackColor = Color.Green OrElse Not ComboBox2.Visible) AndAlso _
                (Label3.BackColor = Color.Green OrElse Not ComboBox3.Visible) AndAlso (Label4.BackColor = Color.Green OrElse Not ComboBox4.Visible) AndAlso _
                (Label5.BackColor = Color.Green OrElse Not ComboBox5.Visible) AndAlso (Label6.BackColor = Color.Green OrElse Not ComboBox6.Visible) Then

                Timer1.Enabled = True

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
            If My.Settings.NovoValorTipoDeJogoTelaEstudoIntervalos = True Then
                Button1.Select()
            Else
                Button3.Select()
            End If
            Desenhar()
            Timer1.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        Try
            If ListBox1.Text = NomeIntervalo Then
                ListBox1.BackColor = Color.Green
                Timer1.Enabled = True
            Else
                ListBox1.BackColor = Color.Red
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        My.Settings.NovoValorTipoDeJogoTelaEstudoIntervalos = False
        Desenhar()
    End Sub
End Class