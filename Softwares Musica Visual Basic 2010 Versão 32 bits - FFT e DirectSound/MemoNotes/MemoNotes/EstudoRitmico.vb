Option Strict On
Option Explicit On

Public Class EstudoRitmico
    Inherits Form 'para mecher no form no modo design é preciso mudar para 'Inherits Form'
    'e depois de ajustado voltar para 'Inherits PerPixelAlphaForm' e clicar para mudar o inherit
    Public TransAmount As Byte = 255
    Dim FaceBit As New Bitmap(My.Resources.TreinamentoRitmico)

    Private lista As New List(Of Keys)
    Private Thread(1) As Thread

    Dim Imagens(3, 4320), ImagemCompasso(1) As Image
    Dim ArrayX(18, 4320), Din As String
    '0 (mão direita) e 1 (mão esquerda) = onde houver "x" ou "xl" significa que há uma nota. O "xl" é para definir a 
    'última nota de um grupo, ou a primeira e última se só existir uma nota no grupo, com 
    'base nos "xl" é possível definir o ponto inicial da ligadura. Se o valor for "p" significa que é uma pausa
    '2 e 3 = onde houver "x" significa que o usuário teclou o ritmo
    '4 = indica a marcação do compasso, onde houver "xx" indica o ínicio do compasso e será adicionado 20 pixeis à 
    'esquerda, para que a primeira nota de cada compasso fique 20 pixeis distante da barra divisória, os demais valores do compasso serão "x"
    '5 e 6 = armazena informações "c" e "e" (certo e errado)
    '7 e 8 = posição inicial e final para desenhar a ligadura, mão direita
    '9 e 10 = posição inicial e final para desenhar a ligadura, mão esquerda
    '11 e 12 = onde houver "l" indica que a nota deve ser exibida mas não tocada, pois é continuidade da nota anterior conectada por ligadura
    '13 e 14 = onde houver "ll" indica que houve mudança de linha, e que a nota anterior inciou uma ligadura que atingirá 
    'a próxima nota na linha abaixo. Então, onde o programa encontrar "ll" desenhará um pequeno arco no ínicio da nova linha de compassos
    '15 e 16 = armazena a intensidade com que as teclas do piano foram pressionadas

    Dim ArrayXX(1, 4320), a, aa, p(2), d, f, g, h, j, l, jj, ii, k, nn, QtdeSubdivisões, Valor(10), AcertosErros(2), NumeroAleatório, NumeroAleatórioLigadura, QtdeCompassos, _
        QtdePixeisCompasso(1), PosiçãoNoArray, PosiçãoLeft, MãoEsqDir, ValorTop(1), CompensaçãoLeft(2), CompensaçãoTop(2), Avanço(2), DistanciaInicialDoCompasso(2), TipoCompasso, _
        LimiteCompasso, AjusteDinâmica(1), PercentualDinâmica, PercentualLigadura, AjusteLigadura, AjusteSwing, QtdePixeisSubdivisão, QtdeIntervalosSubdivisão, BPM, ValorVolumeSom As Integer

    Dim AjusteBPM As Double

    Dim Fonte As New Font("Arial", 9, FontStyle.Regular, GraphicsUnit.Pixel)
    Dim Fonte2 As New Font("Times New Roman", 20, FontStyle.Bold, GraphicsUnit.Pixel)
    Dim Fonte3 As New Font("Arial", 7, FontStyle.Regular, GraphicsUnit.Pixel)
    Dim CorFonte As SolidBrush = New SolidBrush(Color.FromArgb(2, 0, 0))
    Dim CorFonte2 As SolidBrush = New SolidBrush(Color.FromArgb(102, 153, 255))
    Dim PosiçãoExataAtual As SolidBrush = New SolidBrush(Color.FromArgb(0, 0, 0))
    Dim InicioSubdivisão As SolidBrush = New SolidBrush(Color.FromArgb(100, 102, 153, 255))
    Dim LinhaLigadura As Pen = New Pen(Color.FromArgb(0, 0, 0), 1)
    Dim Retangulo As SolidBrush = New SolidBrush(Color.FromArgb(My.Settings.NovaTransparênciaCorLocalizaçãoDosCompassos, My.Settings.NovaCorLocalizaçãoDosCompassos))
    Dim CorMarcaçãoME As SolidBrush = New SolidBrush(My.Settings.NovaCorME)
    Dim CorMarcaçãoMD As SolidBrush = New SolidBrush(My.Settings.NovaCorMD)

    'código para apenas 1 compasso. Começar a programação fazendo apenas 1 compasso, assim fica mais fácil. 
    'Após acertado isso aumenta-se a quantidade de compassos.
    'Fazer primeiramente apenas para as figuras whole, half, quarter e eight

    'whole           1    cada um ocupa 320 pixeis    
    'half               2    cada um ocupa 160 pixeis
    'quarter          4    cada um ocupa 80 pixeis
    'eight             8    cada um ocupa 40 pixeis   43 pixeis no ear master
    'sixteen         16    cada um ocupa 20 pixeis  22 pixeis no ear master
    'thirty-two      32    cada um ocupa 10 pixeis  11 pixeis no ear master
    'sixty-four      64    cada um ocupa 5 pixeis

    Dim c As String  'cada coluna do array corresponde a 1 pixel
    Dim myResourceManager As New Resources.ResourceManager("MemoNotes.Resources", Assembly.GetExecutingAssembly())


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

    Dim ValorSom As GeneralMidiPercussion

    Public microTimer As New MicroLibrary.MicroTimer()


    Public Sub GerarExercícioRitmico()

        Try

            Retangulo = New SolidBrush(Color.FromArgb(My.Settings.NovaTransparênciaCorLocalizaçãoDosCompassos, My.Settings.NovaCorLocalizaçãoDosCompassos))
            CorMarcaçãoME = New SolidBrush(My.Settings.NovaCorME)
            CorMarcaçãoMD = New SolidBrush(My.Settings.NovaCorMD)

            Array.Clear(Imagens, 0, Imagens.Length)
            Array.Clear(ArrayX, 0, ArrayX.Length)
            Array.Clear(ArrayXX, 0, ArrayXX.Length)
            AcertosErros(0) = 0
            AcertosErros(1) = 0
            AcertosErros(2) = -1
            ProgressBar1.Visible = False
            a = 0
            p(0) = -1
            p(1) = -1
            f = 0
            l = 0
            h = 0
            ArmazenaImagensNotasPausasNoArray()
            'Desenhar()

            FaceBit = My.Resources.TreinamentoRitmico
            GerarImagemDeFundo()


            AtualizaRegiãoDasNotas()
            AtualizaRegiãoDoCompasso()
            AtualizaRegiãoPercentualAcertos()


            DistanciaInicialDoCompasso(1) = 0
            DistanciaInicialDoCompasso(2) = 0


            microTimer.Interval = CInt((60000000 / BPM / QtdeIntervalosSubdivisão) * AjusteBPM)
            ' Call micro timer every 1000µs (1ms)
            ' microTimer.IgnoreEventIfLateBy = 500; // Can choose to ignore event if late by Xµs (by default it will try to catch up)
            microTimer.Enabled = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub AtualizaRegiãoDasNotas()
        Dim Rect5 As New Rectangle(60, 135, 1160, 590)
        Me.Invalidate(Rect5)
    End Sub

    Public Sub AtualizaRegiãoDoTempo()
        Dim Rect As New Rectangle(184, 750, 51, 30)
        Me.Invalidate(Rect)
    End Sub

    Private Sub AtualizaRegiãoDoCompasso()
        Dim Rect20 As New Rectangle(111, 65, 64, 50)
        Me.Invalidate(Rect20)
    End Sub

    Private Sub DesenharCCCC()

        Dim gr As Graphics
        Dim FaceBit As New Bitmap(My.Resources.TreinamentoRitmico)
        gr = Graphics.FromImage(FaceBit)

        Dim string_format As New StringFormat
        string_format.FormatFlags = StringFormatFlags.DirectionVertical

        ValorTop(0) = 144
        ValorTop(1) = 155

        CompensaçãoLeft(0) = 0
        CompensaçãoTop(0) = 0

        If p(0) >= 960 AndAlso p(0) < 1920 Then
            CompensaçãoLeft(0) = 960 + 60
            CompensaçãoTop(0) = 121
        ElseIf p(0) >= 1920 AndAlso p(0) < 2880 Then
            CompensaçãoLeft(0) = 1920 + 120
            CompensaçãoTop(0) = 242
        ElseIf p(0) >= 2880 AndAlso p(0) < 3840 Then
            CompensaçãoLeft(0) = 2880 + 180
            CompensaçãoTop(0) = 363
        End If


        If p(0) >= 0 Then
            gr.FillRectangle(Retangulo, p(0) - CompensaçãoLeft(0) + 71 + DistanciaInicialDoCompasso(1), ValorTop(1) + CompensaçãoTop(0), 341, 78)
        End If


        CompensaçãoLeft(0) = 0
        CompensaçãoTop(0) = 0
        CompensaçãoLeft(1) = 0
        CompensaçãoTop(1) = 0
        For i = 0 To 4000

            If ArrayX(4, i) = "xx" Then DistanciaInicialDoCompasso(0) += 20

            If Imagens(0, i) IsNot Nothing Then
                If i >= 960 AndAlso i < 1920 Then
                    CompensaçãoLeft(0) = 960 + 60
                    CompensaçãoTop(0) = 121
                ElseIf i >= 1920 AndAlso i < 2880 Then
                    CompensaçãoLeft(0) = 1920 + 120
                    CompensaçãoTop(0) = 242
                ElseIf i >= 2880 AndAlso i < 3840 Then
                    CompensaçãoLeft(0) = 2880 + 180
                    CompensaçãoTop(0) = 363
                End If

                PosiçãoLeft = i - CompensaçãoLeft(0) + 67 + DistanciaInicialDoCompasso(0)

                gr.DrawImage(Imagens(0, i), PosiçãoLeft, ValorTop(0) + CompensaçãoTop(0), Imagens(0, i).Width, Imagens(0, i).Height)
                'gr.DrawString(CStr(PosiçãoLeft), Fonte, CorFonte, PosiçãoLeft, ValorTop(0) + CompensaçãoTop + 35, string_format)
            End If

            If Imagens(1, i) IsNot Nothing Then
                If i >= 960 AndAlso i < 1920 Then
                    CompensaçãoLeft(1) = 960 + 60
                    CompensaçãoTop(1) = 121
                ElseIf i >= 1920 AndAlso i < 2880 Then
                    CompensaçãoLeft(1) = 1920 + 120
                    CompensaçãoTop(1) = 242
                ElseIf i >= 2880 AndAlso i < 3840 Then
                    CompensaçãoLeft(1) = 2880 + 180
                    CompensaçãoTop(1) = 363
                End If

                PosiçãoLeft = i - CompensaçãoLeft(1) + 67 + DistanciaInicialDoCompasso(0)


                gr.DrawImage(Imagens(1, i), PosiçãoLeft, ValorTop(0) + CompensaçãoTop(1) + 57, Imagens(1, i).Width, Imagens(1, i).Height)
                'gr.DrawString(CStr(PosiçãoLeft), Fonte, CorFonte, PosiçãoLeft, ValorTop(0) + CompensaçãoTop + 35, string_format)
            End If


            PosiçãoLeft = i - CompensaçãoLeft(0) + 69 + DistanciaInicialDoCompasso(0) 'por incrível que pareça, tem que repetir esta linha
            gr.SmoothingMode = SmoothingMode.AntiAlias
            If ArrayX(2, i) = "x" Then
                gr.FillEllipse(Brushes.Black, PosiçãoLeft - 1, ValorTop(0) + CompensaçãoTop(0) + 40, 6, 6)
                gr.FillEllipse(Brushes.Red, PosiçãoLeft, ValorTop(0) + CompensaçãoTop(0) + 41, 4, 4)
                'gr.DrawString(CStr(PosiçãoLeft), Fonte, Brushes.Red, PosiçãoLeft, ValorTop(0) + CompensaçãoTop(0) + 46, string_format)
            End If

            If ArrayX(3, i) = "x" Then
                gr.FillEllipse(Brushes.Black, PosiçãoLeft - 1, ValorTop(0) + CompensaçãoTop(0) + 53, 6, 6)
                gr.FillEllipse(Brushes.DodgerBlue, PosiçãoLeft, ValorTop(0) + CompensaçãoTop(0) + 54, 4, 4)
            End If

            gr.DrawString(CStr(My.Settings.NovoValorTempoInicial), Fonte2, CorFonte, 135, 665)

            gr.SmoothingMode = SmoothingMode.None

            i += 4 'pular campos do array de 5 em cinco
        Next

        'Me.SetBitmap(FaceBit, TransAmount)

    End Sub

    Private Sub DesenharAAA(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint

        Try

            DistanciaInicialDoCompasso(0) = 0

            If g = 0 Then
                e.Graphics.DrawImage(My.Resources.PonteiroMetronomo1, 1086, 64, My.Resources.PonteiroMetronomo1.Width, My.Resources.PonteiroMetronomo1.Height)
            Else
                e.Graphics.DrawImage(My.Resources.PonteiroMetronomo2, 1073, 64, My.Resources.PonteiroMetronomo2.Width, My.Resources.PonteiroMetronomo2.Height)
            End If

            e.Graphics.DrawImage(My.Resources.EfeitoVidro, 1033, 51, My.Resources.EfeitoVidro.Width, My.Resources.EfeitoVidro.Height)


            ValorTop(0) = 144
            ValorTop(1) = 201


            If p(0) >= 0 AndAlso p(0) < LimiteCompasso Then
                CompensaçãoLeft(0) = 20
                CompensaçãoTop(0) = 0
            ElseIf p(0) >= LimiteCompasso AndAlso p(0) < LimiteCompasso * 2 Then
                CompensaçãoLeft(0) = CInt(LimiteCompasso + ((QtdeCompassos / 4) * 20)) + 20
                CompensaçãoTop(0) = 145
            ElseIf p(0) >= LimiteCompasso * 2 AndAlso p(0) < LimiteCompasso * 3 Then
                CompensaçãoLeft(0) = CInt((LimiteCompasso * 2) + ((QtdeCompassos / 4) * 40)) + 20
                CompensaçãoTop(0) = 290
            ElseIf p(0) >= LimiteCompasso * 3 AndAlso p(0) < LimiteCompasso * 4 Then
                CompensaçãoLeft(0) = CInt((LimiteCompasso * 3) + ((QtdeCompassos / 4) * 60)) + 20
                CompensaçãoTop(0) = 435
            End If

            If a < QtdePixeisCompasso(1) * QtdeCompassos Then
                If p(0) >= 0 AndAlso My.Settings.NovoValorLocalizaçãoCompassoAtual = True Then
                    e.Graphics.FillRectangle(Retangulo, p(0) - CompensaçãoLeft(0) + 71 + DistanciaInicialDoCompasso(1), ValorTop(1) + CompensaçãoTop(0), QtdePixeisCompasso(1) + 21, 9)
                End If

                If p(1) >= 0 AndAlso My.Settings.NovoValorLocalizaçãoSubdivisãoCompassoAtual = True Then
                    e.Graphics.FillRectangle(Retangulo, p(1) - CompensaçãoLeft(0) + 91 + DistanciaInicialDoCompasso(1), ValorTop(1) + CompensaçãoTop(0), CInt(QtdePixeisCompasso(1) / QtdeSubdivisões), 9)
                End If

                If p(2) >= 0 AndAlso My.Settings.NovoValorLocalizaçãoMicroSubdivisãoCompassoAtual = True Then
                    e.Graphics.FillRectangle(PosiçãoExataAtual, p(2) - CompensaçãoLeft(0) + 91 + DistanciaInicialDoCompasso(1), ValorTop(1) + CompensaçãoTop(0), 1, 9)
                End If
            End If



            CompensaçãoLeft(0) = 0
            CompensaçãoTop(0) = 0
            CompensaçãoLeft(1) = 0
            CompensaçãoTop(1) = 0

            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias


            For i = 0 To p(2) 'QtdeCompassos * QtdePixeisCompasso(1)

                If ArrayX(4, i) = "xx" Then DistanciaInicialDoCompasso(0) += 20

                If i >= LimiteCompasso AndAlso i < LimiteCompasso * 2 Then
                    CompensaçãoLeft(0) = CInt(LimiteCompasso + ((QtdeCompassos / 4) * 20))
                    CompensaçãoTop(0) = 145
                ElseIf i >= LimiteCompasso * 2 AndAlso i < LimiteCompasso * 3 Then
                    CompensaçãoLeft(0) = CInt((LimiteCompasso * 2) + ((QtdeCompassos / 4) * 40))
                    CompensaçãoTop(0) = 290
                ElseIf i >= LimiteCompasso * 3 AndAlso i < LimiteCompasso * 4 Then
                    CompensaçãoLeft(0) = CInt((LimiteCompasso * 3) + ((QtdeCompassos / 4) * 60))
                    CompensaçãoTop(0) = 435
                End If

                PosiçãoLeft = i - CompensaçãoLeft(0) + 69 + DistanciaInicialDoCompasso(0)

                If ArrayX(2, i) = "x" Then
                    e.Graphics.FillEllipse(Brushes.Black, PosiçãoLeft - 1, ValorTop(0) + CompensaçãoTop(0) + 40, 6, 6)
                    e.Graphics.FillEllipse(CorMarcaçãoMD, PosiçãoLeft, ValorTop(0) + CompensaçãoTop(0) + 41, 4, 4)
                    'e.graphics.DrawString(CStr(PosiçãoLeft), Fonte, Brushes.Red, PosiçãoLeft, ValorTop(0) + CompensaçãoTop(0) + 46, string_format)
                End If

                If ArrayX(3, i) = "x" Then
                    e.Graphics.FillEllipse(Brushes.Black, PosiçãoLeft - 1, ValorTop(0) + CompensaçãoTop(0) + 77, 6, 6)
                    e.Graphics.FillEllipse(CorMarcaçãoME, PosiçãoLeft, ValorTop(0) + CompensaçãoTop(0) + 78, 4, 4)
                End If


                If h = 100 Then 'só executa esta seção após o exercício ter chegado ao fim
                    If ArrayX(5, i) = "" Then

                    ElseIf ArrayX(5, i) = "c" Then
                        e.Graphics.DrawImage(My.Resources.CorretoRitmo, PosiçãoLeft, ValorTop(0) + CompensaçãoTop(0) + 43, 11, 11)
                    ElseIf ArrayX(5, i) = "e" Then
                        e.Graphics.DrawImage(My.Resources.IncorretoRitmo, PosiçãoLeft, ValorTop(0) + CompensaçãoTop(0) + 47, 8, 8)
                    ElseIf ArrayX(5, i) = "ed" Then
                        e.Graphics.DrawImage(My.Resources.IncorretoRitmoDinamica, PosiçãoLeft, ValorTop(0) + CompensaçãoTop(0) + 47, 14, 8)
                        e.Graphics.DrawString(ArrayX(15, i), Fonte3, CorFonte, PosiçãoLeft, ValorTop(0) + CompensaçãoTop(0) + 55)
                    End If


                    If ArrayX(6, i) = "" Then

                    ElseIf ArrayX(6, i) = "c" Then
                        e.Graphics.DrawImage(My.Resources.CorretoRitmo, PosiçãoLeft, ValorTop(0) + CompensaçãoTop(0) + 66, 11, 11)
                    ElseIf ArrayX(6, i) = "e" Then
                        e.Graphics.DrawImage(My.Resources.IncorretoRitmo, PosiçãoLeft, ValorTop(0) + CompensaçãoTop(0) + 70, 8, 8)
                    ElseIf ArrayX(6, i) = "ed" Then
                        e.Graphics.DrawImage(My.Resources.IncorretoRitmoDinamica, PosiçãoLeft, ValorTop(0) + CompensaçãoTop(0) + 70, 14, 8)
                        e.Graphics.DrawString(ArrayX(16, i), Fonte3, CorFonte, PosiçãoLeft, ValorTop(0) + CompensaçãoTop(0) + 65)
                    End If
                End If



                i += 4 'pular campos do array de 5 em cinco
            Next


            e.Graphics.DrawString(CStr(My.Settings.NovoValorTempoInicial), Fonte2, CorFonte, 185, 757)
            If AcertosErros(2) >= 0 Then
                e.Graphics.DrawString(AcertosErros(2) & "%", Fonte2, CorFonte, 700, 757)
                ProgressBar1.Visible = True
                ProgressBar1.Value = AcertosErros(2)
            End If

            e.Graphics.SmoothingMode = SmoothingMode.None

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Public Sub GerarImagemDeFundo()

        Try

            Dim gr As Graphics = Graphics.FromImage(FaceBit)

            DistanciaInicialDoCompasso(0) = 0

            gr.DrawImage(ImagemCompasso(0), 127, 60, ImagemCompasso(0).Width, ImagemCompasso(0).Height)
            gr.DrawImage(My.Resources.EfeitoVidro, 88, 51, My.Resources.EfeitoVidro.Width, My.Resources.EfeitoVidro.Height)

            gr.DrawImage(My.Resources.BarraCompasso2, 74, 155, My.Resources.BarraCompasso2.Width, My.Resources.BarraCompasso2.Height)
            gr.DrawImage(My.Resources.BarraCompasso2, 74, 206, My.Resources.BarraCompasso2.Width, My.Resources.BarraCompasso2.Height)

            gr.DrawImage(ImagemCompasso(1), 71, 155, ImagemCompasso(1).Width, ImagemCompasso(1).Height)
            gr.DrawImage(ImagemCompasso(1), 71, 206, ImagemCompasso(1).Width, ImagemCompasso(1).Height)
            gr.DrawImage(ImagemCompasso(1), 71, 300, ImagemCompasso(1).Width, ImagemCompasso(1).Height)
            gr.DrawImage(ImagemCompasso(1), 71, 351, ImagemCompasso(1).Width, ImagemCompasso(1).Height)
            gr.DrawImage(ImagemCompasso(1), 71, 445, ImagemCompasso(1).Width, ImagemCompasso(1).Height)
            gr.DrawImage(ImagemCompasso(1), 71, 496, ImagemCompasso(1).Width, ImagemCompasso(1).Height)
            gr.DrawImage(ImagemCompasso(1), 71, 590, ImagemCompasso(1).Width, ImagemCompasso(1).Height)
            gr.DrawImage(ImagemCompasso(1), 71, 641, ImagemCompasso(1).Width, ImagemCompasso(1).Height)

            gr.DrawImage(My.Resources.BarraCompasso2, CInt(((QtdeCompassos / 4) * QtdePixeisCompasso(1)) + ((QtdeCompassos / 4) * 20) + 68), 590, My.Resources.BarraCompasso2.Width, My.Resources.BarraCompasso2.Height)
            gr.DrawImage(My.Resources.BarraCompasso2, CInt(((QtdeCompassos / 4) * QtdePixeisCompasso(1)) + ((QtdeCompassos / 4) * 20) + 68), 641, My.Resources.BarraCompasso2.Width, My.Resources.BarraCompasso2.Height)



            Dim string_format As New StringFormat
            string_format.FormatFlags = StringFormatFlags.DirectionVertical

            ValorTop(0) = 144
            ValorTop(1) = 201

            If My.Settings.NovoValorFiguraRitmicas(29) = "True" OrElse My.Settings.NovoValorFiguraRitmicas(30) = "True" OrElse My.Settings.NovoValorFiguraRitmicas(31) = "True" OrElse My.Settings.NovoValorFiguraRitmicas(32) = "True" Then
                AjusteDinâmica(0) = 8
            Else
                AjusteDinâmica(0) = -3
            End If
            If My.Settings.NovoValorFiguraRitmicas(13) = "True" OrElse My.Settings.NovoValorFiguraRitmicas(14) = "True" OrElse My.Settings.NovoValorFiguraRitmicas(15) = "True" OrElse My.Settings.NovoValorFiguraRitmicas(16) = "True" Then
                AjusteDinâmica(1) = 122
            Else
                AjusteDinâmica(1) = 111
            End If



            CompensaçãoLeft(0) = 0
            CompensaçãoTop(0) = 0
            CompensaçãoLeft(1) = 0
            CompensaçãoTop(1) = 0


            For i = 0 To QtdeCompassos * QtdePixeisCompasso(1)

                If ArrayX(4, i) = "xx" Then DistanciaInicialDoCompasso(0) += 20 : nn = 0

                If My.Settings.NovoValorMãoDireita = True Then
                    If Imagens(0, i) IsNot Nothing Then
                        If i >= LimiteCompasso AndAlso i < LimiteCompasso * 2 Then
                            CompensaçãoLeft(0) = CInt(LimiteCompasso + ((QtdeCompassos / 4) * 20))
                            CompensaçãoTop(0) = 145
                        ElseIf i >= LimiteCompasso * 2 AndAlso i < LimiteCompasso * 3 Then
                            CompensaçãoLeft(0) = CInt((LimiteCompasso * 2) + ((QtdeCompassos / 4) * 40))
                            CompensaçãoTop(0) = 290
                        ElseIf i >= LimiteCompasso * 3 AndAlso i < LimiteCompasso * 4 Then
                            CompensaçãoLeft(0) = CInt((LimiteCompasso * 3) + ((QtdeCompassos / 4) * 60))
                            CompensaçãoTop(0) = 435
                        End If

                        PosiçãoLeft = i - CompensaçãoLeft(0) + 67 + DistanciaInicialDoCompasso(0)

                        gr.DrawImage(Imagens(0, i), PosiçãoLeft, ValorTop(0) + CompensaçãoTop(0), Imagens(0, i).Width, Imagens(0, i).Height)
                        'gr.DrawString(CStr(i), Fonte, CorFonte, PosiçãoLeft, ValorTop(0) + CompensaçãoTop(0) + 45, string_format)

                    End If


                    If My.Settings.NovoValorDinamicaMD = True AndAlso Imagens(2, i) IsNot Nothing Then
                        gr.DrawImage(Imagens(2, i), PosiçãoLeft, ValorTop(0) + CompensaçãoTop(0) - AjusteDinâmica(0), Imagens(2, i).Width, Imagens(2, i).Height)
                    End If

                    gr.SmoothingMode = SmoothingMode.AntiAlias
                    If My.Settings.NovoValorLigadurasMD = True AndAlso ArrayX(7, i) <> "" Then
                        gr.DrawArc(LinhaLigadura, CInt(ArrayX(7, i)) - CompensaçãoLeft(0) + 74 + DistanciaInicialDoCompasso(0), ValorTop(0) + CompensaçãoTop(0) + 35, CInt(ArrayX(8, i)) - CInt(ArrayX(7, i)) - 6, 10, 180, -180)
                    End If
                    If My.Settings.NovoValorLigadurasMD = True AndAlso ArrayX(13, i) = "ll" Then
                        gr.DrawArc(LinhaLigadura, 53, ValorTop(0) + CompensaçãoTop(0) + 35, 34, 10, 140, -140)
                    End If
                    gr.SmoothingMode = SmoothingMode.None
                End If


                If My.Settings.NovoValorMãoEsquerda = True Then
                    If Imagens(1, i) IsNot Nothing Then
                        If i >= LimiteCompasso AndAlso i < LimiteCompasso * 2 Then
                            CompensaçãoLeft(1) = CInt(LimiteCompasso + ((QtdeCompassos / 4) * 20))
                            CompensaçãoTop(1) = 145
                        ElseIf i >= LimiteCompasso * 2 AndAlso i < LimiteCompasso * 3 Then
                            CompensaçãoLeft(1) = CInt((LimiteCompasso * 2) + ((QtdeCompassos / 4) * 40))
                            CompensaçãoTop(1) = 290
                        ElseIf i >= LimiteCompasso * 3 AndAlso i < LimiteCompasso * 4 Then
                            CompensaçãoLeft(1) = CInt((LimiteCompasso * 3) + ((QtdeCompassos / 4) * 60))
                            CompensaçãoTop(1) = 435
                        End If

                        PosiçãoLeft = i - CompensaçãoLeft(1) + 67 + DistanciaInicialDoCompasso(0)


                        gr.DrawImage(Imagens(1, i), PosiçãoLeft, ValorTop(0) + CompensaçãoTop(1) + 81, Imagens(1, i).Width, Imagens(1, i).Height)
                        'gr.DrawString(CStr(PosiçãoLeft), Fonte, CorFonte, PosiçãoLeft, ValorTop(0) + CompensaçãoTop + 35, string_format)
                    End If


                    If My.Settings.NovoValorDinamicaME = True AndAlso Imagens(3, i) IsNot Nothing Then
                        gr.DrawImage(Imagens(3, i), PosiçãoLeft - 5, ValorTop(0) + CompensaçãoTop(1) + AjusteDinâmica(1), Imagens(3, i).Width, Imagens(3, i).Height)
                    End If

                    gr.SmoothingMode = SmoothingMode.AntiAlias
                    If My.Settings.NovoValorLigadurasME = True AndAlso ArrayX(9, i) <> "" Then
                        gr.DrawArc(LinhaLigadura, CInt(ArrayX(9, i)) - CompensaçãoLeft(1) + 73 + DistanciaInicialDoCompasso(0), ValorTop(0) + CompensaçãoTop(1) + 78, CInt(ArrayX(10, i)) - CInt(ArrayX(9, i)) - 6, 10, 180, 180)
                    End If
                    If My.Settings.NovoValorLigadurasME = True AndAlso ArrayX(14, i) = "ll" Then
                        gr.DrawArc(LinhaLigadura, 53, ValorTop(0) + CompensaçãoTop(1) + 78, 34, 10, 220, 140)
                    End If
                    gr.SmoothingMode = SmoothingMode.None
                End If



                PosiçãoLeft = i - CompensaçãoLeft(0) + 69 + DistanciaInicialDoCompasso(0)  'por incrível que pareça, tem que repetir esta linha



                If ArrayX(4, i) = "x" Or ArrayX(4, i) = "xx" Then
                    nn += 1
                    gr.FillRectangle(InicioSubdivisão, PosiçãoLeft + 2, ValorTop(0) + CompensaçãoTop(0) + 60, 1, 3)
                    gr.DrawString(CStr(nn), Fonte3, CorFonte2, PosiçãoLeft + 3, ValorTop(0) + CompensaçãoTop(0) + 58)
                End If


                i += 9 'pular campos do array de 10 em cinco
            Next

            Me.BackgroundImage = FaceBit

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub ArmazenaImagensNotasPausasNoArray()

        Try

            GerarTipoCompasso()

            If My.Settings.NovoValorMãoDireita = True Then
                PosiçãoNoArray = 0 : MãoEsqDir = 0 : c = ""
                If My.Settings.NovoValorSwingMD = True Then AjusteSwing = 20 Else AjusteSwing = 0
                GeraNotas()
                If My.Settings.NovoValorLigadurasMD = True Then GerarLigaduras()
                If My.Settings.NovoValorDinamicaMD = True Then GerarDinâmicaDasNotas() 'Dinâmica deve ser gerada depois das ligaduras, pois a aplicação de ligaduras fará com que algumas notas, mesmo que exibidas, não devam ser tocadas
            End If
            If My.Settings.NovoValorMãoEsquerda = True Then
                PosiçãoNoArray = 0 : MãoEsqDir = 1 : c = "N"
                If My.Settings.NovoValorSwingME = True Then AjusteSwing = 20 Else AjusteSwing = 0
                GeraNotas()
                If My.Settings.NovoValorLigadurasME = True Then GerarLigaduras()
                If My.Settings.NovoValorDinamicaME = True Then GerarDinâmicaDasNotas()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub GerarDinâmicaDasNotas()

        Try

            For i = 0 To QtdeCompassos * QtdePixeisCompasso(1)
                If ArrayX(MãoEsqDir, i) = "x" OrElse ArrayX(MãoEsqDir, i) = "xl" Then
                    ii = 1000

                    Do While ii > 19 'array é zero based, então isso abrange 20 números, 0 à 19
                        Dim randomNumber(0) As Byte
                        Dim Gen As New Security.Cryptography.RNGCryptoServiceProvider()
                        Gen.GetBytes(randomNumber)
                        ii = Int(Convert.ToInt32(randomNumber(0)))
                    Loop

                    If MãoEsqDir = 0 Then
                        PercentualDinâmica = My.Settings.NovoValorPercentualDinamicaMD
                    Else
                        PercentualDinâmica = My.Settings.NovoValorPercentualDinamicaME
                    End If

                    If Percentual(ii, PercentualDinâmica) = "x" Then 'ii dará o valor da linha, Acidentes.Value dará o valor da coluna do array
                        ii = 1000
                        Do While ii < 1 OrElse ii > 8 'define a força da dinâmica
                            Dim randomNumber(0) As Byte
                            Dim Gen As New Security.Cryptography.RNGCryptoServiceProvider()
                            Gen.GetBytes(randomNumber)
                            ii = Int(Convert.ToInt32(randomNumber(0)))

                            'mão direita
                            'mão direita
                            If MãoEsqDir = 0 Then
                                If ii = 1 AndAlso My.Settings.NovoValorForçaDinâmica(7) = "False" Then 'FFFMD
                                    ii = 100
                                ElseIf ii = 2 AndAlso My.Settings.NovoValorForçaDinâmica(6) = "False" Then 'FFMD
                                    ii = 100
                                ElseIf ii = 3 AndAlso My.Settings.NovoValorForçaDinâmica(5) = "False" Then 'FMD
                                    ii = 100
                                ElseIf ii = 4 AndAlso My.Settings.NovoValorForçaDinâmica(4) = "False" Then 'MFMD
                                    ii = 100
                                ElseIf ii = 5 AndAlso My.Settings.NovoValorForçaDinâmica(3) = "False" Then 'MPMD
                                    ii = 100
                                ElseIf ii = 6 AndAlso My.Settings.NovoValorForçaDinâmica(2) = "False" Then 'PMD
                                    ii = 100
                                ElseIf ii = 7 AndAlso My.Settings.NovoValorForçaDinâmica(1) = "False" Then 'PPMD
                                    ii = 100
                                ElseIf ii = 8 AndAlso My.Settings.NovoValorForçaDinâmica(0) = "False" Then 'PPPMD
                                    ii = 100
                                End If
                            End If
                            'mão esquerda
                            If MãoEsqDir = 1 Then
                                If ii = 1 AndAlso My.Settings.NovoValorForçaDinâmica(15) = "False" Then 'FFFME
                                    ii = 100
                                ElseIf ii = 2 AndAlso My.Settings.NovoValorForçaDinâmica(14) = "False" Then 'FFME
                                    ii = 100
                                ElseIf ii = 3 AndAlso My.Settings.NovoValorForçaDinâmica(13) = "False" Then 'FME
                                    ii = 100
                                ElseIf ii = 4 AndAlso My.Settings.NovoValorForçaDinâmica(12) = "False" Then 'MFME
                                    ii = 100
                                ElseIf ii = 5 AndAlso My.Settings.NovoValorForçaDinâmica(11) = "False" Then 'MPME
                                    ii = 100
                                ElseIf ii = 6 AndAlso My.Settings.NovoValorForçaDinâmica(10) = "False" Then 'PME
                                    ii = 100
                                ElseIf ii = 7 AndAlso My.Settings.NovoValorForçaDinâmica(9) = "False" Then 'PPME
                                    ii = 100
                                ElseIf ii = 8 AndAlso My.Settings.NovoValorForçaDinâmica(8) = "False" Then 'PPPME
                                    ii = 100
                                End If
                            End If

                        Loop

                        If MãoEsqDir = 0 Then
                            If ii = 1 Then
                                Imagens(2, i) = My.Resources.Dfff : ArrayX(17, i) = "fff"
                            ElseIf ii = 2 Then
                                Imagens(2, i) = My.Resources.Dff : ArrayX(17, i) = "ff"
                            ElseIf ii = 3 Then
                                Imagens(2, i) = My.Resources.Df : ArrayX(17, i) = "f"
                            ElseIf ii = 4 Then
                                Imagens(2, i) = My.Resources.Dmf : ArrayX(17, i) = "mf"
                            ElseIf ii = 5 Then
                                Imagens(2, i) = My.Resources.Dmp : ArrayX(17, i) = "mp"
                            ElseIf ii = 6 Then
                                Imagens(2, i) = My.Resources.Dp : ArrayX(17, i) = "p"
                            ElseIf ii = 7 Then
                                Imagens(2, i) = My.Resources.Dpp : ArrayX(17, i) = "pp"
                            ElseIf ii = 8 Then
                                Imagens(2, i) = My.Resources.Dppp : ArrayX(17, i) = "ppp"
                            End If
                        Else
                            If ii = 1 Then
                                Imagens(3, i) = My.Resources.Dfff2 : ArrayX(18, i) = "fff"
                            ElseIf ii = 2 Then
                                Imagens(3, i) = My.Resources.Dff2 : ArrayX(18, i) = "ff"
                            ElseIf ii = 3 Then
                                Imagens(3, i) = My.Resources.Df2 : ArrayX(18, i) = "f"
                            ElseIf ii = 4 Then
                                Imagens(3, i) = My.Resources.Dmf2 : ArrayX(18, i) = "mf"
                            ElseIf ii = 5 Then
                                Imagens(3, i) = My.Resources.Dmp2 : ArrayX(18, i) = "mp"
                            ElseIf ii = 6 Then
                                Imagens(3, i) = My.Resources.Dp2 : ArrayX(18, i) = "p"
                            ElseIf ii = 7 Then
                                Imagens(3, i) = My.Resources.Dpp2 : ArrayX(18, i) = "pp"
                            ElseIf ii = 8 Then
                                Imagens(3, i) = My.Resources.Dppp2 : ArrayX(18, i) = "ppp"
                            End If
                        End If
                    End If
                End If
                Avanço(1) = i : AvançoRápido() : i += Avanço(2)
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub GerarLigaduras()

        Try

            For i = 0 To QtdeCompassos * QtdePixeisCompasso(1)
                If ArrayX(MãoEsqDir, i) = "xl" Then
                    ii = 1000

                    Do While ii > 19 'array é zero based, então isso abrange 20 números, 0 à 19
                        Dim randomNumber(0) As Byte
                        Dim Gen As New Security.Cryptography.RNGCryptoServiceProvider()
                        Gen.GetBytes(randomNumber)
                        ii = Int(Convert.ToInt32(randomNumber(0)))
                    Loop
                    If MãoEsqDir = 0 Then
                        PercentualLigadura = My.Settings.NovoValorPercentualLigadurasMD
                        k = 0
                    Else
                        PercentualLigadura = My.Settings.NovoValorPercentualLigadurasME
                        k = 2
                    End If

                    If Percentual(ii, PercentualLigadura) = "x" Then 'ii dará o valor da linha, Acidentes.Value dará o valor da coluna do array
                        AjusteLigadura = 0
                        For iii = i + 5 To i + 1000
                            If ArrayX(4, iii) = "xx" Then AjusteLigadura += 20 'ajusta a ligadura que atravessa de um compasso a outro
                            If ArrayX(MãoEsqDir, iii) = "p" Then
                                iii = 50000 'cai fora do loop
                            ElseIf ArrayX(MãoEsqDir, iii) = "x" OrElse ArrayX(MãoEsqDir, iii) = "xl" Then
                                ArrayX(k + 7, i) = CStr(i) 'insere o valor da posição incial, na linha 7 ou na linha 9 do array
                                ArrayX(k + 8, i) = CStr(iii + AjusteLigadura) 'insere o valor da posição final, na linha 8 ou na linha 10 do array
                                ArrayX(MãoEsqDir + 11, iii) = "l" 'indica que esta nota deverá ser exibida, mas não tocada

                                If (i < LimiteCompasso AndAlso iii >= LimiteCompasso) OrElse _
                                    (i < LimiteCompasso * 2 AndAlso iii >= LimiteCompasso * 2) OrElse _
                                    (i < LimiteCompasso * 3 AndAlso iii >= LimiteCompasso * 3) OrElse _
                                    (i < LimiteCompasso * 4 AndAlso iii >= LimiteCompasso * 4) Then
                                    ArrayX(MãoEsqDir + 13, iii) = "ll"
                                End If


                                iii = 50000 'cai fora do loop
                            End If
                            iii += 4 'pular campos do array de 5 em cinco
                            If iii > 4320 Then iii = 50000 'cai fora do loop
                        Next
                    End If
                End If
                Avanço(1) = i : AvançoRápido() : i += Avanço(2)
            Next

            For i = 0 To QtdeCompassos * QtdePixeisCompasso(1)
                If ArrayX(MãoEsqDir + 11, i) = "l" Then ArrayX(MãoEsqDir, i) = "" 'fará com que a nota não seja tocada, afinal ela é uma prolongação da nota anterior
                Avanço(1) = i : AvançoRápido() : i += Avanço(2)
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub AvançoRápido()

        Try

            If ArrayXX(MãoEsqDir, Avanço(1)) = 60 OrElse ArrayXX(MãoEsqDir, Avanço(1)) = 63 Then
                Avanço(2) = 599
            ElseIf ArrayXX(MãoEsqDir, Avanço(1)) = 57 OrElse ArrayXX(MãoEsqDir, Avanço(1)) = 59 Then
                Avanço(2) = 559
            ElseIf ArrayXX(MãoEsqDir, Avanço(1)) = 56 OrElse ArrayXX(MãoEsqDir, Avanço(1)) = 58 Then
                Avanço(2) = 479
            ElseIf ArrayXX(MãoEsqDir, Avanço(1)) = 1 OrElse ArrayXX(MãoEsqDir, Avanço(1)) = 31 Then
                Avanço(2) = 319
            ElseIf ArrayXX(MãoEsqDir, Avanço(1)) = 61 OrElse ArrayXX(MãoEsqDir, Avanço(1)) = 64 Then
                Avanço(2) = 299
            ElseIf ArrayXX(MãoEsqDir, Avanço(1)) = 50 OrElse ArrayXX(MãoEsqDir, Avanço(1)) = 51 Then
                Avanço(2) = 279
            ElseIf ArrayXX(MãoEsqDir, Avanço(1)) = 42 OrElse ArrayXX(MãoEsqDir, Avanço(1)) = 43 Then
                Avanço(2) = 239
            ElseIf ArrayXX(MãoEsqDir, Avanço(1)) = 2 OrElse ArrayXX(MãoEsqDir, Avanço(1)) = 30 Then
                Avanço(2) = 159
            ElseIf ArrayXX(MãoEsqDir, Avanço(1)) = 62 OrElse ArrayXX(MãoEsqDir, Avanço(1)) = 65 Then
                Avanço(2) = 149
            ElseIf ArrayXX(MãoEsqDir, Avanço(1)) = 52 OrElse ArrayXX(MãoEsqDir, Avanço(1)) = 53 Then
                Avanço(2) = 139
            ElseIf ArrayXX(MãoEsqDir, Avanço(1)) = 44 OrElse ArrayXX(MãoEsqDir, Avanço(1)) = 45 Then
                Avanço(2) = 119
            ElseIf ArrayXX(MãoEsqDir, Avanço(1)) = 3 OrElse ArrayXX(MãoEsqDir, Avanço(1)) = 13 Then
                Avanço(2) = 79
            ElseIf ArrayXX(MãoEsqDir, Avanço(1)) = 54 OrElse ArrayXX(MãoEsqDir, Avanço(1)) = 55 Then
                Avanço(2) = 69
            ElseIf ArrayXX(MãoEsqDir, Avanço(1)) = 46 OrElse ArrayXX(MãoEsqDir, Avanço(1)) = 47 Then
                Avanço(2) = 59
            ElseIf ArrayXX(MãoEsqDir, Avanço(1)) = 4 OrElse ArrayXX(MãoEsqDir, Avanço(1)) = 26 Then
                Avanço(2) = 39
            ElseIf ArrayXX(MãoEsqDir, Avanço(1)) = 66 OrElse ArrayXX(MãoEsqDir, Avanço(1)) = 67 Then
                Avanço(2) = 29
            ElseIf ArrayXX(MãoEsqDir, Avanço(1)) = 12 OrElse ArrayXX(MãoEsqDir, Avanço(1)) = 27 Then
                Avanço(2) = 19
            ElseIf ArrayXX(MãoEsqDir, Avanço(1)) = 28 OrElse ArrayXX(MãoEsqDir, Avanço(1)) = 29 Then
                Avanço(2) = 9
            Else
                Avanço(2) = 9 'pular campos do array de 9 em 9
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub GerarTipoCompasso()

        Try

            TipoCompasso = 100
            Do While TipoCompasso < 1 OrElse TipoCompasso > 17
                Dim randomNumber(0) As Byte
                Dim Gen As New Security.Cryptography.RNGCryptoServiceProvider()
                Gen.GetBytes(randomNumber)
                TipoCompasso = Int(Convert.ToInt32(randomNumber(0)))

                If TipoCompasso > 0 AndAlso TipoCompasso < 18 AndAlso My.Settings.NovoValorCompassos(TipoCompasso - 1) = "False" Then TipoCompasso = 100

            Loop

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub GeraNotas()

        Try

            If TipoCompasso = 1 Then '2/2
                ImagemCompasso(0) = My.Resources._2_2
                ImagemCompasso(1) = My.Resources.BarraCompasso1
                QtdeCompassos = 12
                QtdePixeisCompasso(1) = 320
                QtdeSubdivisões = 2
                AjusteBPM = 2
            ElseIf TipoCompasso = 2 Then '3/2
                ImagemCompasso(0) = My.Resources._3_2
                ImagemCompasso(1) = My.Resources.BarraCompasso3
                QtdeCompassos = 8
                QtdePixeisCompasso(1) = 480
                QtdeSubdivisões = 3
                AjusteBPM = 2
            ElseIf TipoCompasso = 3 Then '4/2
                ImagemCompasso(0) = My.Resources._4_2
                ImagemCompasso(1) = My.Resources.BarraCompasso13
                QtdeCompassos = 4
                QtdePixeisCompasso(1) = 640
                QtdeSubdivisões = 4
                AjusteBPM = 2
            ElseIf TipoCompasso = 4 Then '2/4
                ImagemCompasso(0) = My.Resources._2_4
                ImagemCompasso(1) = My.Resources.BarraCompasso5
                QtdeCompassos = 24
                QtdePixeisCompasso(1) = 160
                QtdeSubdivisões = 2
                AjusteBPM = 1
            ElseIf TipoCompasso = 5 Then '3/4
                ImagemCompasso(0) = My.Resources._3_4
                ImagemCompasso(1) = My.Resources.BarraCompasso4
                QtdeCompassos = 16
                QtdePixeisCompasso(1) = 240
                QtdeSubdivisões = 3
                AjusteBPM = 1
            ElseIf TipoCompasso = 6 Then '4/4
                ImagemCompasso(0) = My.Resources._4_4
                ImagemCompasso(1) = My.Resources.BarraCompasso1
                QtdeCompassos = 12
                QtdePixeisCompasso(1) = 320
                QtdeSubdivisões = 4
                AjusteBPM = 1
            ElseIf TipoCompasso = 7 Then '5/4
                ImagemCompasso(0) = My.Resources._5_4
                ImagemCompasso(1) = My.Resources.BarraCompasso10
                QtdeCompassos = 8
                QtdePixeisCompasso(1) = 400
                QtdeSubdivisões = 5
                AjusteBPM = 1
            ElseIf TipoCompasso = 8 Then '6/4
                ImagemCompasso(0) = My.Resources._6_4
                ImagemCompasso(1) = My.Resources.BarraCompasso3
                QtdeCompassos = 8
                QtdePixeisCompasso(1) = 480
                QtdeSubdivisões = 6
                AjusteBPM = 1
            ElseIf TipoCompasso = 9 Then '9/4
                ImagemCompasso(0) = My.Resources._9_4
                ImagemCompasso(1) = My.Resources.BarraCompasso11
                QtdeCompassos = 4
                QtdePixeisCompasso(1) = 720
                QtdeSubdivisões = 9
                AjusteBPM = 1
            ElseIf TipoCompasso = 10 Then '12/4
                ImagemCompasso(0) = My.Resources._12_4
                ImagemCompasso(1) = My.Resources.BarraCompasso12
                QtdeCompassos = 4
                QtdePixeisCompasso(1) = 960
                QtdeSubdivisões = 12
                AjusteBPM = 1
            ElseIf TipoCompasso = 11 Then '3/8
                ImagemCompasso(0) = My.Resources._3_8
                ImagemCompasso(1) = My.Resources.BarraCompasso9
                QtdeCompassos = 32
                QtdePixeisCompasso(1) = 120
                QtdeSubdivisões = 3
                AjusteBPM = 0.5
            ElseIf TipoCompasso = 12 Then '4/8
                ImagemCompasso(0) = My.Resources._4_8
                ImagemCompasso(1) = My.Resources.BarraCompasso5
                QtdeCompassos = 24
                QtdePixeisCompasso(1) = 160
                QtdeSubdivisões = 4
                AjusteBPM = 0.5
            ElseIf TipoCompasso = 13 Then '5/8
                ImagemCompasso(0) = My.Resources._5_8
                ImagemCompasso(1) = My.Resources.BarraCompasso6
                QtdeCompassos = 20
                QtdePixeisCompasso(1) = 200
                QtdeSubdivisões = 5
                AjusteBPM = 0.5
            ElseIf TipoCompasso = 14 Then '6/8
                ImagemCompasso(0) = My.Resources._6_8
                ImagemCompasso(1) = My.Resources.BarraCompasso4
                QtdeCompassos = 16
                QtdePixeisCompasso(1) = 240
                QtdeSubdivisões = 6
                AjusteBPM = 0.5
            ElseIf TipoCompasso = 15 Then '7/8
                ImagemCompasso(0) = My.Resources._7_8
                ImagemCompasso(1) = My.Resources.BarraCompasso7
                QtdeCompassos = 12
                QtdePixeisCompasso(1) = 280
                QtdeSubdivisões = 7
                AjusteBPM = 0.5
            ElseIf TipoCompasso = 16 Then '9/8
                ImagemCompasso(0) = My.Resources._9_8
                ImagemCompasso(1) = My.Resources.BarraCompasso8
                QtdeCompassos = 12
                QtdePixeisCompasso(1) = 360
                QtdeSubdivisões = 9
                AjusteBPM = 0.5
            ElseIf TipoCompasso = 17 Then '12/8
                ImagemCompasso(0) = My.Resources._12_8
                ImagemCompasso(1) = My.Resources.BarraCompasso3
                QtdeCompassos = 8
                QtdePixeisCompasso(1) = 480
                QtdeSubdivisões = 12
                AjusteBPM = 0.5
            End If

            BPM = CInt(My.Settings.NovoValorTempoInicial)
            QtdePixeisSubdivisão = CInt(QtdePixeisCompasso(1) / QtdeSubdivisões)
            QtdeIntervalosSubdivisão = CInt(QtdePixeisSubdivisão / 5)

            LimiteCompasso = CInt((QtdeCompassos / 4) * QtdePixeisCompasso(1))


            For i = 1 To QtdeCompassos

                ArrayX(4, PosiçãoNoArray) = "xx"

                If TipoCompasso = 1 Then '2/2
                    ArrayX(4, PosiçãoNoArray + 160) = "x"
                ElseIf TipoCompasso = 2 Then '3/2
                    ArrayX(4, PosiçãoNoArray + 160) = "x"
                    ArrayX(4, PosiçãoNoArray + 320) = "x"
                ElseIf TipoCompasso = 3 Then '4/2
                    ArrayX(4, PosiçãoNoArray + 160) = "x"
                    ArrayX(4, PosiçãoNoArray + 320) = "x"
                    ArrayX(4, PosiçãoNoArray + 480) = "x"
                ElseIf TipoCompasso = 4 Then '2/4
                    ArrayX(4, PosiçãoNoArray + 80) = "x"
                ElseIf TipoCompasso = 5 Then '3/4
                    ArrayX(4, PosiçãoNoArray + 80) = "x"
                    ArrayX(4, PosiçãoNoArray + 160) = "x"
                ElseIf TipoCompasso = 6 Then '4/4
                    ArrayX(4, PosiçãoNoArray + 80) = "x"
                    ArrayX(4, PosiçãoNoArray + 160) = "x"
                    ArrayX(4, PosiçãoNoArray + 240) = "x"
                ElseIf TipoCompasso = 7 Then '5/4
                    ArrayX(4, PosiçãoNoArray + 80) = "x"
                    ArrayX(4, PosiçãoNoArray + 160) = "x"
                    ArrayX(4, PosiçãoNoArray + 240) = "x"
                    ArrayX(4, PosiçãoNoArray + 320) = "x"
                ElseIf TipoCompasso = 8 Then '6/4
                    ArrayX(4, PosiçãoNoArray + 80) = "x"
                    ArrayX(4, PosiçãoNoArray + 160) = "x"
                    ArrayX(4, PosiçãoNoArray + 240) = "x"
                    ArrayX(4, PosiçãoNoArray + 320) = "x"
                    ArrayX(4, PosiçãoNoArray + 400) = "x"
                ElseIf TipoCompasso = 9 Then '9/4
                    ArrayX(4, PosiçãoNoArray + 80) = "x"
                    ArrayX(4, PosiçãoNoArray + 160) = "x"
                    ArrayX(4, PosiçãoNoArray + 240) = "x"
                    ArrayX(4, PosiçãoNoArray + 320) = "x"
                    ArrayX(4, PosiçãoNoArray + 400) = "x"
                    ArrayX(4, PosiçãoNoArray + 480) = "x"
                    ArrayX(4, PosiçãoNoArray + 560) = "x"
                    ArrayX(4, PosiçãoNoArray + 640) = "x"
                ElseIf TipoCompasso = 10 Then '12/4
                    ArrayX(4, PosiçãoNoArray + 80) = "x"
                    ArrayX(4, PosiçãoNoArray + 160) = "x"
                    ArrayX(4, PosiçãoNoArray + 240) = "x"
                    ArrayX(4, PosiçãoNoArray + 320) = "x"
                    ArrayX(4, PosiçãoNoArray + 400) = "x"
                    ArrayX(4, PosiçãoNoArray + 480) = "x"
                    ArrayX(4, PosiçãoNoArray + 560) = "x"
                    ArrayX(4, PosiçãoNoArray + 640) = "x"
                    ArrayX(4, PosiçãoNoArray + 720) = "x"
                    ArrayX(4, PosiçãoNoArray + 800) = "x"
                    ArrayX(4, PosiçãoNoArray + 880) = "x"
                ElseIf TipoCompasso = 11 Then '3/8
                    ArrayX(4, PosiçãoNoArray + 40) = "x"
                    ArrayX(4, PosiçãoNoArray + 80) = "x"
                ElseIf TipoCompasso = 12 Then '4/8
                    ArrayX(4, PosiçãoNoArray + 40) = "x"
                    ArrayX(4, PosiçãoNoArray + 80) = "x"
                    ArrayX(4, PosiçãoNoArray + 120) = "x"
                ElseIf TipoCompasso = 13 Then '5/8
                    ArrayX(4, PosiçãoNoArray + 40) = "x"
                    ArrayX(4, PosiçãoNoArray + 80) = "x"
                    ArrayX(4, PosiçãoNoArray + 120) = "x"
                    ArrayX(4, PosiçãoNoArray + 160) = "x"
                ElseIf TipoCompasso = 14 Then '6/8
                    ArrayX(4, PosiçãoNoArray + 40) = "x"
                    ArrayX(4, PosiçãoNoArray + 80) = "x"
                    ArrayX(4, PosiçãoNoArray + 120) = "x"
                    ArrayX(4, PosiçãoNoArray + 160) = "x"
                    ArrayX(4, PosiçãoNoArray + 200) = "x"
                ElseIf TipoCompasso = 15 Then '7/8
                    ArrayX(4, PosiçãoNoArray + 40) = "x"
                    ArrayX(4, PosiçãoNoArray + 80) = "x"
                    ArrayX(4, PosiçãoNoArray + 120) = "x"
                    ArrayX(4, PosiçãoNoArray + 160) = "x"
                    ArrayX(4, PosiçãoNoArray + 200) = "x"
                    ArrayX(4, PosiçãoNoArray + 240) = "x"
                ElseIf TipoCompasso = 16 Then '9/8
                    ArrayX(4, PosiçãoNoArray + 40) = "x"
                    ArrayX(4, PosiçãoNoArray + 80) = "x"
                    ArrayX(4, PosiçãoNoArray + 120) = "x"
                    ArrayX(4, PosiçãoNoArray + 160) = "x"
                    ArrayX(4, PosiçãoNoArray + 200) = "x"
                    ArrayX(4, PosiçãoNoArray + 240) = "x"
                    ArrayX(4, PosiçãoNoArray + 280) = "x"
                    ArrayX(4, PosiçãoNoArray + 320) = "x"
                ElseIf TipoCompasso = 17 Then '12/8
                    ArrayX(4, PosiçãoNoArray + 40) = "x"
                    ArrayX(4, PosiçãoNoArray + 80) = "x"
                    ArrayX(4, PosiçãoNoArray + 120) = "x"
                    ArrayX(4, PosiçãoNoArray + 160) = "x"
                    ArrayX(4, PosiçãoNoArray + 200) = "x"
                    ArrayX(4, PosiçãoNoArray + 240) = "x"
                    ArrayX(4, PosiçãoNoArray + 280) = "x"
                    ArrayX(4, PosiçãoNoArray + 320) = "x"
                    ArrayX(4, PosiçãoNoArray + 360) = "x"
                    ArrayX(4, PosiçãoNoArray + 400) = "x"
                    ArrayX(4, PosiçãoNoArray + 440) = "x"
                End If


                QtdePixeisCompasso(0) = QtdePixeisCompasso(1)
                Do While QtdePixeisCompasso(0) > 0
                    NumeroAleatório = 100
                    Do While NumeroAleatório < 1 OrElse NumeroAleatório > 73 'preciso ver que número usar. Vou ter que fazer uma tabela com todas combinações, usando exemplos do earmaster 


                        ' Create a byte array to hold the random value.
                        Dim randomNumber(0) As Byte

                        ' Create a new instance of the RNGCryptoServiceProvider.
                        Dim Gen As New Security.Cryptography.RNGCryptoServiceProvider()

                        ' Fill the array with a random value.
                        Gen.GetBytes(randomNumber)

                        ' Convert the byte to an integer value to make the modulus operation easier.
                        NumeroAleatório = Int(Convert.ToInt32(randomNumber(0)))

                        If PosiçãoNoArray = 0 AndAlso (NumeroAleatório = 13 OrElse NumeroAleatório = 26 OrElse NumeroAleatório = 27 OrElse NumeroAleatório = 29 OrElse NumeroAleatório = 30 OrElse NumeroAleatório = 31 _
                                    OrElse NumeroAleatório = 10 OrElse NumeroAleatório = 18 OrElse NumeroAleatório = 22 OrElse NumeroAleatório = 23 OrElse NumeroAleatório = 33 OrElse NumeroAleatório = 34 _
                                    OrElse NumeroAleatório = 38 OrElse NumeroAleatório = 39 OrElse NumeroAleatório = 43 OrElse NumeroAleatório = 45 OrElse NumeroAleatório = 47 OrElse NumeroAleatório = 51 _
                                    OrElse NumeroAleatório = 53 OrElse NumeroAleatório = 55 OrElse NumeroAleatório = 58 OrElse NumeroAleatório = 59 OrElse NumeroAleatório = 63 OrElse NumeroAleatório = 64 _
                                    OrElse NumeroAleatório = 65 OrElse NumeroAleatório = 67 OrElse NumeroAleatório = 69 OrElse NumeroAleatório = 72) Then
                            'fará com que a primeira figura do primeiro compasso não seja uma pausa
                            NumeroAleatório = 100
                        End If

                        'mão esquerda
                        If MãoEsqDir = 1 Then
                            If NumeroAleatório = 1 AndAlso My.Settings.NovoValorFiguraRitmicas(1) = "False" Then
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 2 AndAlso My.Settings.NovoValorFiguraRitmicas(2) = "False" Then
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 3 AndAlso My.Settings.NovoValorFiguraRitmicas(3) = "False" Then
                                NumeroAleatório = 100
                            ElseIf (NumeroAleatório = 4 OrElse NumeroAleatório = 5 OrElse NumeroAleatório = 6) AndAlso My.Settings.NovoValorFiguraRitmicas(4) = "False" Then
                                NumeroAleatório = 100
                            ElseIf (NumeroAleatório = 12 OrElse NumeroAleatório = 7 OrElse NumeroAleatório = 25) AndAlso My.Settings.NovoValorFiguraRitmicas(5) = "False" Then
                                NumeroAleatório = 100
                            ElseIf (NumeroAleatório = 14 OrElse NumeroAleatório = 15 OrElse NumeroAleatório = 16) AndAlso (My.Settings.NovoValorFiguraRitmicas(5) = "False" OrElse My.Settings.NovoValorFiguraRitmicas(4) = "False") Then
                                NumeroAleatório = 100
                            ElseIf (NumeroAleatório = 28 OrElse NumeroAleatório = 73) AndAlso My.Settings.NovoValorFiguraRitmicas(6) = "False" Then
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 31 AndAlso My.Settings.NovoValorFiguraRitmicas(7) = "False" Then
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 30 AndAlso My.Settings.NovoValorFiguraRitmicas(8) = "False" Then
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 13 AndAlso My.Settings.NovoValorFiguraRitmicas(9) = "False" Then
                                NumeroAleatório = 100
                            ElseIf (NumeroAleatório = 26) AndAlso My.Settings.NovoValorFiguraRitmicas(10) = "False" Then
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 24 AndAlso (My.Settings.NovoValorFiguraRitmicas(10) = "False" OrElse My.Settings.NovoValorFiguraRitmicas(5) = "False") Then
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 27 AndAlso My.Settings.NovoValorFiguraRitmicas(11) = "False" Then
                                NumeroAleatório = 100
                            ElseIf (NumeroAleatório = 8 OrElse NumeroAleatório = 9 OrElse NumeroAleatório = 10 OrElse NumeroAleatório = 11 OrElse NumeroAleatório = 22 OrElse NumeroAleatório = 23) AndAlso (My.Settings.NovoValorFiguraRitmicas(11) = "False" OrElse My.Settings.NovoValorFiguraRitmicas(5) = "False") Then
                                NumeroAleatório = 100
                            ElseIf (NumeroAleatório = 17 OrElse NumeroAleatório = 18 OrElse NumeroAleatório = 19 OrElse NumeroAleatório = 20 OrElse NumeroAleatório = 21) AndAlso (My.Settings.NovoValorFiguraRitmicas(11) = "False" OrElse My.Settings.NovoValorFiguraRitmicas(4) = "False" OrElse My.Settings.NovoValorFiguraRitmicas(5) = "False") Then
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 29 AndAlso My.Settings.NovoValorFiguraRitmicas(12) = "False" Then
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 32 AndAlso My.Settings.NovoValorFiguraRitmicas(13) = "False" Then
                                NumeroAleatório = 100
                            ElseIf (NumeroAleatório = 33 OrElse NumeroAleatório = 34 OrElse NumeroAleatório = 35 OrElse NumeroAleatório = 36) AndAlso (My.Settings.NovoValorFiguraRitmicas(13) = "False" OrElse My.Settings.NovoValorFiguraRitmicas(8) = "False") Then
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 37 AndAlso My.Settings.NovoValorFiguraRitmicas(14) = "False" Then
                                NumeroAleatório = 100
                            ElseIf (NumeroAleatório = 38 OrElse NumeroAleatório = 39 OrElse NumeroAleatório = 40 OrElse NumeroAleatório = 41) AndAlso (My.Settings.NovoValorFiguraRitmicas(14) = "False" OrElse My.Settings.NovoValorFiguraRitmicas(9) = "False") Then
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 68 AndAlso My.Settings.NovoValorFiguraRitmicas(15) = "False" Then
                                NumeroAleatório = 100
                            ElseIf (NumeroAleatório = 69 OrElse NumeroAleatório = 70 OrElse NumeroAleatório = 71 OrElse NumeroAleatório = 72) AndAlso (My.Settings.NovoValorFiguraRitmicas(15) = "False" OrElse My.Settings.NovoValorFiguraRitmicas(10) = "False") Then
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 56 AndAlso (My.Settings.NovoValorPontoAumentoME = False OrElse My.Settings.NovoValorFiguraRitmicas(1) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(2) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(24) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(3) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(9) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(4) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(10) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(5) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(11) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(6) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(12) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 58 AndAlso (My.Settings.NovoValorPontoAumentoME = False OrElse My.Settings.NovoValorFiguraRitmicas(7) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(2) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(24) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(3) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(9) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(4) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(10) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(5) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(11) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(6) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(12) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 42 AndAlso (My.Settings.NovoValorPontoAumentoME = False OrElse My.Settings.NovoValorFiguraRitmicas(2) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(3) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(9) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(4) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(10) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(5) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(11) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(6) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(12) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 43 AndAlso (My.Settings.NovoValorPontoAumentoME = False OrElse My.Settings.NovoValorFiguraRitmicas(8) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(9) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(3) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(10) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(4) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(11) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(5) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(12) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(6) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 44 AndAlso (My.Settings.NovoValorPontoAumentoME = False OrElse My.Settings.NovoValorFiguraRitmicas(3) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(4) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(10) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(5) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(11) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(6) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(12) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 45 AndAlso (My.Settings.NovoValorPontoAumentoME = False OrElse My.Settings.NovoValorFiguraRitmicas(9) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(10) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(4) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(11) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(5) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(12) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(6) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 46 AndAlso (My.Settings.NovoValorPontoAumentoME = False OrElse My.Settings.NovoValorFiguraRitmicas(4) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(5) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(11) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(6) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(12) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 47 AndAlso (My.Settings.NovoValorPontoAumentoME = False OrElse My.Settings.NovoValorFiguraRitmicas(10) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(11) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(5) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(12) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(6) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 66 AndAlso (My.Settings.NovoValorPontoAumentoME = False OrElse My.Settings.NovoValorFiguraRitmicas(5) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(6) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(12) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 67 AndAlso (My.Settings.NovoValorPontoAumentoME = False OrElse My.Settings.NovoValorFiguraRitmicas(11) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(12) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(6) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf (NumeroAleatório = 48 OrElse NumeroAleatório = 49) AndAlso (My.Settings.NovoValorPontoAumentoME = False OrElse My.Settings.NovoValorFiguraRitmicas(5) = "False" OrElse My.Settings.NovoValorFiguraRitmicas(4) = "False") Then
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 57 AndAlso (My.Settings.NovoValorDuploPontoAumentoME = False OrElse My.Settings.NovoValorFiguraRitmicas(1) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(3) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(9) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(4) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(10) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(5) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(11) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(6) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(12) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 59 AndAlso (My.Settings.NovoValorDuploPontoAumentoME = False OrElse My.Settings.NovoValorFiguraRitmicas(7) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(3) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(9) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(10) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(4) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(11) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(5) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(12) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(6) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 50 AndAlso (My.Settings.NovoValorDuploPontoAumentoME = False OrElse My.Settings.NovoValorFiguraRitmicas(2) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(4) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(10) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(5) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(11) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(6) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(12) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 51 AndAlso (My.Settings.NovoValorDuploPontoAumentoME = False OrElse My.Settings.NovoValorFiguraRitmicas(8) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(10) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(4) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(11) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(5) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(12) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(6) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 52 AndAlso (My.Settings.NovoValorDuploPontoAumentoME = False OrElse My.Settings.NovoValorFiguraRitmicas(3) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(5) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(11) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(6) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(12) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 53 AndAlso (My.Settings.NovoValorDuploPontoAumentoME = False OrElse My.Settings.NovoValorFiguraRitmicas(9) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(11) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(5) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(12) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(6) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 54 AndAlso (My.Settings.NovoValorDuploPontoAumentoME = False OrElse My.Settings.NovoValorFiguraRitmicas(4) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(6) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(12) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 55 AndAlso (My.Settings.NovoValorDuploPontoAumentoME = False OrElse My.Settings.NovoValorFiguraRitmicas(10) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(12) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(6) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 60 AndAlso (My.Settings.NovoValorTriploPontoAumentoME = False OrElse My.Settings.NovoValorFiguraRitmicas(1) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(4) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(10) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(5) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(11) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(6) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(12) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 63 AndAlso (My.Settings.NovoValorTriploPontoAumentoME = False OrElse My.Settings.NovoValorFiguraRitmicas(7) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(10) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(4) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(11) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(5) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(12) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(6) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 61 AndAlso (My.Settings.NovoValorTriploPontoAumentoME = False OrElse My.Settings.NovoValorFiguraRitmicas(2) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(5) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(11) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(6) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(12) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 64 AndAlso (My.Settings.NovoValorTriploPontoAumentoME = False OrElse My.Settings.NovoValorFiguraRitmicas(8) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(11) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(5) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(12) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(6) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 62 AndAlso (My.Settings.NovoValorTriploPontoAumentoME = False OrElse My.Settings.NovoValorFiguraRitmicas(3) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(6) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(12) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 65 AndAlso (My.Settings.NovoValorTriploPontoAumentoME = False OrElse My.Settings.NovoValorFiguraRitmicas(9) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(12) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(6) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            End If
                        End If


                        'mão direita
                        If MãoEsqDir = 0 Then
                            If NumeroAleatório = 1 AndAlso My.Settings.NovoValorFiguraRitmicas(17) = "False" Then
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 2 AndAlso My.Settings.NovoValorFiguraRitmicas(18) = "False" Then
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 3 AndAlso My.Settings.NovoValorFiguraRitmicas(19) = "False" Then
                                NumeroAleatório = 100
                            ElseIf (NumeroAleatório = 4 OrElse NumeroAleatório = 5 OrElse NumeroAleatório = 6) AndAlso My.Settings.NovoValorFiguraRitmicas(20) = "False" Then
                                NumeroAleatório = 100
                            ElseIf (NumeroAleatório = 12 OrElse NumeroAleatório = 7 OrElse NumeroAleatório = 25) AndAlso My.Settings.NovoValorFiguraRitmicas(21) = "False" Then
                                NumeroAleatório = 100
                            ElseIf (NumeroAleatório = 14 OrElse NumeroAleatório = 15 OrElse NumeroAleatório = 16) AndAlso (My.Settings.NovoValorFiguraRitmicas(21) = "False" OrElse My.Settings.NovoValorFiguraRitmicas(20) = "False") Then
                                NumeroAleatório = 100
                            ElseIf (NumeroAleatório = 28 OrElse NumeroAleatório = 73) AndAlso My.Settings.NovoValorFiguraRitmicas(22) = "False" Then
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 31 AndAlso My.Settings.NovoValorFiguraRitmicas(23) = "False" Then
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 30 AndAlso My.Settings.NovoValorFiguraRitmicas(24) = "False" Then
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 13 AndAlso My.Settings.NovoValorFiguraRitmicas(25) = "False" Then
                                NumeroAleatório = 100
                            ElseIf (NumeroAleatório = 26) AndAlso My.Settings.NovoValorFiguraRitmicas(26) = "False" Then
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 24 AndAlso (My.Settings.NovoValorFiguraRitmicas(26) = "False" OrElse My.Settings.NovoValorFiguraRitmicas(21) = "False") Then
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 27 AndAlso My.Settings.NovoValorFiguraRitmicas(27) = "False" Then
                                NumeroAleatório = 100
                            ElseIf (NumeroAleatório = 8 OrElse NumeroAleatório = 9 OrElse NumeroAleatório = 10 OrElse NumeroAleatório = 11 OrElse NumeroAleatório = 22 OrElse NumeroAleatório = 23) AndAlso (My.Settings.NovoValorFiguraRitmicas(27) = "False" OrElse My.Settings.NovoValorFiguraRitmicas(21) = "False") Then
                                NumeroAleatório = 100
                            ElseIf (NumeroAleatório = 17 OrElse NumeroAleatório = 18 OrElse NumeroAleatório = 19 OrElse NumeroAleatório = 20 OrElse NumeroAleatório = 21) AndAlso (My.Settings.NovoValorFiguraRitmicas(27) = "False" OrElse My.Settings.NovoValorFiguraRitmicas(21) = "False" OrElse My.Settings.NovoValorFiguraRitmicas(20) = "False") Then
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 29 AndAlso My.Settings.NovoValorFiguraRitmicas(28) = "False" Then
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 32 AndAlso My.Settings.NovoValorFiguraRitmicas(29) = "False" Then
                                NumeroAleatório = 100
                            ElseIf (NumeroAleatório = 33 OrElse NumeroAleatório = 34 OrElse NumeroAleatório = 35 OrElse NumeroAleatório = 36) AndAlso (My.Settings.NovoValorFiguraRitmicas(29) = "False" OrElse My.Settings.NovoValorFiguraRitmicas(24) = "False") Then
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 37 AndAlso My.Settings.NovoValorFiguraRitmicas(30) = "False" Then
                                NumeroAleatório = 100
                            ElseIf (NumeroAleatório = 38 OrElse NumeroAleatório = 39 OrElse NumeroAleatório = 40 OrElse NumeroAleatório = 41) AndAlso (My.Settings.NovoValorFiguraRitmicas(30) = "False" OrElse My.Settings.NovoValorFiguraRitmicas(25) = "False") Then
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 68 AndAlso My.Settings.NovoValorFiguraRitmicas(31) = "False" Then
                                NumeroAleatório = 100
                            ElseIf (NumeroAleatório = 69 OrElse NumeroAleatório = 70 OrElse NumeroAleatório = 71 OrElse NumeroAleatório = 72) AndAlso (My.Settings.NovoValorFiguraRitmicas(31) = "False" OrElse My.Settings.NovoValorFiguraRitmicas(26) = "False") Then
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 56 AndAlso (My.Settings.NovoValorPontoAumentoMD = False OrElse My.Settings.NovoValorFiguraRitmicas(17) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(18) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(24) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(19) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(25) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(20) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(26) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(21) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(27) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(22) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(28) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 58 AndAlso (My.Settings.NovoValorPontoAumentoMD = False OrElse My.Settings.NovoValorFiguraRitmicas(23) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(18) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(24) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(25) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(19) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(26) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(20) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(27) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(21) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(28) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(22) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 42 AndAlso (My.Settings.NovoValorPontoAumentoMD = False OrElse My.Settings.NovoValorFiguraRitmicas(18) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(19) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(25) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(20) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(26) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(21) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(27) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(22) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(28) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 43 AndAlso (My.Settings.NovoValorPontoAumentoMD = False OrElse My.Settings.NovoValorFiguraRitmicas(24) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(25) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(19) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(26) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(20) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(27) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(21) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(28) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(22) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 44 AndAlso (My.Settings.NovoValorPontoAumentoMD = False OrElse My.Settings.NovoValorFiguraRitmicas(19) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(20) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(26) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(21) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(27) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(22) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(28) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 45 AndAlso (My.Settings.NovoValorPontoAumentoMD = False OrElse My.Settings.NovoValorFiguraRitmicas(25) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(26) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(20) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(27) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(21) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(28) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(22) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 46 AndAlso (My.Settings.NovoValorPontoAumentoMD = False OrElse My.Settings.NovoValorFiguraRitmicas(20) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(21) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(27) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(22) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(28) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 47 AndAlso (My.Settings.NovoValorPontoAumentoMD = False OrElse My.Settings.NovoValorFiguraRitmicas(26) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(27) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(21) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(28) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(22) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 66 AndAlso (My.Settings.NovoValorPontoAumentoMD = False OrElse My.Settings.NovoValorFiguraRitmicas(21) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(22) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(28) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 67 AndAlso (My.Settings.NovoValorPontoAumentoMD = False OrElse My.Settings.NovoValorFiguraRitmicas(27) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(28) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(22) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf (NumeroAleatório = 48 OrElse NumeroAleatório = 49) AndAlso (My.Settings.NovoValorPontoAumentoMD = False OrElse My.Settings.NovoValorFiguraRitmicas(21) = "False" OrElse My.Settings.NovoValorFiguraRitmicas(20) = "False") Then
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 57 AndAlso (My.Settings.NovoValorDuploPontoAumentoMD = False OrElse My.Settings.NovoValorFiguraRitmicas(17) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(19) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(25) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(20) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(26) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(21) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(27) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(22) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(28) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 59 AndAlso (My.Settings.NovoValorDuploPontoAumentoMD = False OrElse My.Settings.NovoValorFiguraRitmicas(23) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(19) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(25) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(26) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(20) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(27) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(21) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(28) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(22) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 50 AndAlso (My.Settings.NovoValorDuploPontoAumentoMD = False OrElse My.Settings.NovoValorFiguraRitmicas(18) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(20) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(26) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(21) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(27) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(22) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(28) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 51 AndAlso (My.Settings.NovoValorDuploPontoAumentoMD = False OrElse My.Settings.NovoValorFiguraRitmicas(24) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(26) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(20) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(27) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(21) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(28) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(22) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 52 AndAlso (My.Settings.NovoValorDuploPontoAumentoMD = False OrElse My.Settings.NovoValorFiguraRitmicas(19) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(21) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(27) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(22) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(28) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 53 AndAlso (My.Settings.NovoValorDuploPontoAumentoMD = False OrElse My.Settings.NovoValorFiguraRitmicas(25) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(27) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(21) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(28) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(22) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 54 AndAlso (My.Settings.NovoValorDuploPontoAumentoMD = False OrElse My.Settings.NovoValorFiguraRitmicas(20) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(22) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(28) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 55 AndAlso (My.Settings.NovoValorDuploPontoAumentoMD = False OrElse My.Settings.NovoValorFiguraRitmicas(26) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(28) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(22) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 60 AndAlso (My.Settings.NovoValorTriploPontoAumentoMD = False OrElse My.Settings.NovoValorFiguraRitmicas(17) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(20) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(26) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(21) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(27) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(22) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(28) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 63 AndAlso (My.Settings.NovoValorTriploPontoAumentoMD = False OrElse My.Settings.NovoValorFiguraRitmicas(23) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(26) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(20) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(27) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(21) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(28) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(22) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 61 AndAlso (My.Settings.NovoValorTriploPontoAumentoMD = False OrElse My.Settings.NovoValorFiguraRitmicas(18) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(21) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(27) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(22) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(28) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 64 AndAlso (My.Settings.NovoValorTriploPontoAumentoMD = False OrElse My.Settings.NovoValorFiguraRitmicas(24) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(27) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(21) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(28) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(22) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 62 AndAlso (My.Settings.NovoValorTriploPontoAumentoMD = False OrElse My.Settings.NovoValorFiguraRitmicas(19) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(22) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(28) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            ElseIf NumeroAleatório = 65 AndAlso (My.Settings.NovoValorTriploPontoAumentoMD = False OrElse My.Settings.NovoValorFiguraRitmicas(25) = "False" OrElse (My.Settings.NovoValorFiguraRitmicas(28) = "False" AndAlso My.Settings.NovoValorFiguraRitmicas(22) = "False")) Then 'notas pontuadas exigem no minimo a nota que vem depois... ex.: half pontuada exige que no mínimo quarter esteja marcado, caso contrário o compasso não poderá ser fechado
                                NumeroAleatório = 100
                            End If
                        End If
                    Loop

                    ' If My.Settings.NovoValorLigadurasMD = True Then
                    'NumeroAleatórioLigadura = 100
                    ''Do While NumeroAleatórioLigadura < 1 OrElse NumeroAleatórioLigadura > 10
                    'Dim randomNumber(0) As Byte
                    'Dim Gen As New Security.Cryptography.RNGCryptoServiceProvider()
                    'Gen.GetBytes(randomNumber)
                    ' NumeroAleatórioLigadura = Int(Convert.ToInt32(randomNumber(0)))
                    'Loop
                    ' End If

                    ArrayX(MãoEsqDir, PosiçãoNoArray) = "p" 'Valor inicial será "p", indicando que é uma pausa. Caso o valor gere uma nota em
                    'vez de uma pausa, então o valor será alterado. Isso evita com que esta linha tenha que ser
                    'repetida em cada tipo diferente de pausa

                    If NumeroAleatório = 1 AndAlso QtdePixeisCompasso(0) - 320 >= 0 Then
                        Avanço(0) = 320
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "xl" 'xl indica a última nota do grupo, ou a primeira se for nota única. Com base neste valor "xl" é criado as ligaduras
                        GeraNotas2()
                    ElseIf NumeroAleatório = 2 AndAlso QtdePixeisCompasso(0) - 160 >= 0 Then
                        Avanço(0) = 160
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 3 AndAlso QtdePixeisCompasso(0) - 80 >= 0 Then
                        Avanço(0) = 80
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 4 AndAlso QtdePixeisCompasso(0) - 40 >= 0 Then
                        Avanço(0) = 40
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 5 AndAlso QtdePixeisCompasso(0) - 80 >= 0 Then
                        Avanço(0) = 80
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "x"
                        'If (MãoEsqDir = 0 AndAlso SwingMD.checked) OrElse (MãoEsqDir = 1 AndAlso SwingME.checked) Then 'se swing estiver ativado
                        'ArrayX(MãoEsqDir, PosiçãoNoArray + 60) = "xl"
                        'Else 'se swing não estiver ativado
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 40 + AjusteSwing) = "xl"
                        'End If
                        GeraNotas2()
                    ElseIf NumeroAleatório = 6 AndAlso QtdePixeisCompasso(0) - 160 >= 0 Then
                        Avanço(0) = 160
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 40 + AjusteSwing) = "x"
                        'If (MãoEsqDir = 0 AndAlso SwingMD.checked) OrElse (MãoEsqDir = 1 AndAlso SwingME.checked) Then 'se swing estiver ativado
                        'ArrayX(MãoEsqDir, PosiçãoNoArray + 60) = "x"
                        'ArrayX(MãoEsqDir, PosiçãoNoArray + 140) = "xl"
                        'Else 'se swing não estiver ativado
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 80) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 120 + AjusteSwing) = "xl"
                        'End If
                        GeraNotas2()
                    ElseIf NumeroAleatório = 7 AndAlso QtdePixeisCompasso(0) - 80 >= 0 Then
                        Avanço(0) = 80
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 20) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 40) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 60) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 8 AndAlso QtdePixeisCompasso(0) - 80 >= 0 Then
                        Avanço(0) = 80
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 20) = "p"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 40) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 60) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 9 AndAlso QtdePixeisCompasso(0) - 80 >= 0 Then
                        Avanço(0) = 80
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 20) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 40) = "p"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 60) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 10 AndAlso QtdePixeisCompasso(0) - 80 >= 0 Then
                        Avanço(0) = 80
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 20) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 40) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 60) = "x"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 11 AndAlso QtdePixeisCompasso(0) - 80 >= 0 Then
                        Avanço(0) = 80
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 20) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 40) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 60) = "p"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 12 AndAlso QtdePixeisCompasso(0) - 20 >= 0 Then
                        Avanço(0) = 20
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 13 AndAlso QtdePixeisCompasso(0) - 80 >= 0 Then
                        Avanço(0) = 80
                        GeraNotas2()
                    ElseIf NumeroAleatório = 14 AndAlso QtdePixeisCompasso(0) - 80 >= 0 Then
                        Avanço(0) = 80
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 20) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 40) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 15 AndAlso QtdePixeisCompasso(0) - 80 >= 0 Then
                        Avanço(0) = 80
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 20) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 60) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 16 AndAlso QtdePixeisCompasso(0) - 80 >= 0 Then
                        Avanço(0) = 80
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 40) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 60) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 17 AndAlso QtdePixeisCompasso(0) - 80 >= 0 Then
                        Avanço(0) = 80
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 40) = "p"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 60) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 18 AndAlso QtdePixeisCompasso(0) - 80 >= 0 Then
                        Avanço(0) = 80
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 20) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 60) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 19 AndAlso QtdePixeisCompasso(0) - 80 >= 0 Then
                        Avanço(0) = 80
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 20) = "p"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 40) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 20 AndAlso QtdePixeisCompasso(0) - 80 >= 0 Then
                        Avanço(0) = 80
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 40) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 60) = "p"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 21 AndAlso QtdePixeisCompasso(0) - 80 >= 0 Then
                        Avanço(0) = 80
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 20) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 60) = "p"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 22 AndAlso QtdePixeisCompasso(0) - 80 >= 0 Then
                        Avanço(0) = 80
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 20) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 40) = "p"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 60) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 23 AndAlso QtdePixeisCompasso(0) - 80 >= 0 Then
                        Avanço(0) = 80
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 20) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 40) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 60) = "p"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 24 AndAlso QtdePixeisCompasso(0) - 80 >= 0 Then
                        Avanço(0) = 80
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 20) = "p"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 60) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 25 AndAlso QtdePixeisCompasso(0) - 40 >= 0 Then
                        Avanço(0) = 40
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 20) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 26 AndAlso QtdePixeisCompasso(0) - 40 >= 0 Then
                        Avanço(0) = 40
                        GeraNotas2()
                    ElseIf NumeroAleatório = 27 AndAlso QtdePixeisCompasso(0) - 20 >= 0 Then
                        Avanço(0) = 20
                        GeraNotas2()
                    ElseIf NumeroAleatório = 28 AndAlso QtdePixeisCompasso(0) - 10 >= 0 Then
                        Avanço(0) = 10
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 29 AndAlso QtdePixeisCompasso(0) - 10 >= 0 Then
                        Avanço(0) = 10
                        GeraNotas2()
                    ElseIf NumeroAleatório = 30 AndAlso QtdePixeisCompasso(0) - 160 >= 0 Then
                        Avanço(0) = 160
                        GeraNotas2()
                    ElseIf NumeroAleatório = 31 AndAlso QtdePixeisCompasso(0) - 320 >= 0 Then
                        Avanço(0) = 320
                        GeraNotas2()
                    ElseIf NumeroAleatório = 32 AndAlso QtdePixeisCompasso(0) - 320 >= 0 Then
                        Avanço(0) = 320
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 105) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 210) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 33 AndAlso QtdePixeisCompasso(0) - 320 >= 0 Then
                        Avanço(0) = 320
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 105) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 210) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 34 AndAlso QtdePixeisCompasso(0) - 320 >= 0 Then
                        Avanço(0) = 320
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 105) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 210) = "p"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 35 AndAlso QtdePixeisCompasso(0) - 320 >= 0 Then
                        Avanço(0) = 320
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 105) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 210) = "p"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 36 AndAlso QtdePixeisCompasso(0) - 320 >= 0 Then
                        Avanço(0) = 320
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 105) = "p"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 210) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 37 AndAlso QtdePixeisCompasso(0) - 160 >= 0 Then
                        Avanço(0) = 160
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 50) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 105) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 38 AndAlso QtdePixeisCompasso(0) - 160 >= 0 Then
                        Avanço(0) = 160
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 50) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 105) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 39 AndAlso QtdePixeisCompasso(0) - 160 >= 0 Then
                        Avanço(0) = 160
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 50) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 105) = "p"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 40 AndAlso QtdePixeisCompasso(0) - 160 >= 0 Then
                        Avanço(0) = 160
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 50) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 105) = "p"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 41 AndAlso QtdePixeisCompasso(0) - 160 >= 0 Then
                        Avanço(0) = 160
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 50) = "p"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 105) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 42 AndAlso QtdePixeisCompasso(0) - 240 >= 0 Then
                        Avanço(0) = 240
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 43 AndAlso QtdePixeisCompasso(0) - 240 >= 0 Then
                        Avanço(0) = 240
                        GeraNotas2()
                    ElseIf NumeroAleatório = 44 AndAlso QtdePixeisCompasso(0) - 120 >= 0 Then
                        Avanço(0) = 120
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 45 AndAlso QtdePixeisCompasso(0) - 120 >= 0 Then
                        Avanço(0) = 120
                        GeraNotas2()
                    ElseIf NumeroAleatório = 46 AndAlso QtdePixeisCompasso(0) - 60 >= 0 Then
                        Avanço(0) = 60
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 47 AndAlso QtdePixeisCompasso(0) - 60 >= 0 Then
                        Avanço(0) = 60
                        GeraNotas2()
                    ElseIf NumeroAleatório = 48 AndAlso QtdePixeisCompasso(0) - 80 >= 0 Then
                        Avanço(0) = 80
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 60) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 49 AndAlso QtdePixeisCompasso(0) - 80 >= 0 Then
                        Avanço(0) = 80
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 20) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 50 AndAlso QtdePixeisCompasso(0) - 280 >= 0 Then
                        Avanço(0) = 280
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 51 AndAlso QtdePixeisCompasso(0) - 280 >= 0 Then
                        Avanço(0) = 280
                        GeraNotas2()
                    ElseIf NumeroAleatório = 52 AndAlso QtdePixeisCompasso(0) - 140 >= 0 Then
                        Avanço(0) = 140
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 53 AndAlso QtdePixeisCompasso(0) - 140 >= 0 Then
                        Avanço(0) = 140
                        GeraNotas2()
                    ElseIf NumeroAleatório = 54 AndAlso QtdePixeisCompasso(0) - 70 >= 0 Then
                        Avanço(0) = 70
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 55 AndAlso QtdePixeisCompasso(0) - 70 >= 0 Then
                        Avanço(0) = 70
                        GeraNotas2()
                    ElseIf NumeroAleatório = 56 AndAlso QtdePixeisCompasso(0) - 480 >= 0 Then
                        Avanço(0) = 480
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 57 AndAlso QtdePixeisCompasso(0) - 560 >= 0 Then
                        Avanço(0) = 560
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 58 AndAlso QtdePixeisCompasso(0) - 480 >= 0 Then
                        Avanço(0) = 480
                        GeraNotas2()
                    ElseIf NumeroAleatório = 59 AndAlso QtdePixeisCompasso(0) - 560 >= 0 Then
                        Avanço(0) = 560
                        GeraNotas2()
                    ElseIf NumeroAleatório = 60 AndAlso QtdePixeisCompasso(0) - 600 >= 0 Then
                        Avanço(0) = 600
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 61 AndAlso QtdePixeisCompasso(0) - 300 >= 0 Then
                        Avanço(0) = 300
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 62 AndAlso QtdePixeisCompasso(0) - 150 >= 0 Then
                        Avanço(0) = 150
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 63 AndAlso QtdePixeisCompasso(0) - 600 >= 0 Then
                        Avanço(0) = 600
                        GeraNotas2()
                    ElseIf NumeroAleatório = 64 AndAlso QtdePixeisCompasso(0) - 300 >= 0 Then
                        Avanço(0) = 300
                        GeraNotas2()
                    ElseIf NumeroAleatório = 65 AndAlso QtdePixeisCompasso(0) - 150 >= 0 Then
                        Avanço(0) = 150
                        GeraNotas2()
                    ElseIf NumeroAleatório = 66 AndAlso QtdePixeisCompasso(0) - 30 >= 0 Then
                        Avanço(0) = 30
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "x"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 67 AndAlso QtdePixeisCompasso(0) - 30 >= 0 Then
                        Avanço(0) = 30
                        GeraNotas2()
                    ElseIf NumeroAleatório = 68 AndAlso QtdePixeisCompasso(0) - 80 >= 0 Then
                        Avanço(0) = 80
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 25) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 50) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 69 AndAlso QtdePixeisCompasso(0) - 80 >= 0 Then
                        Avanço(0) = 80
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 25) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 50) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 70 AndAlso QtdePixeisCompasso(0) - 80 >= 0 Then
                        Avanço(0) = 80
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 25) = "p"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 50) = "xl"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 71 AndAlso QtdePixeisCompasso(0) - 80 >= 0 Then
                        Avanço(0) = 80
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 25) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 50) = "p"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 72 AndAlso QtdePixeisCompasso(0) - 80 >= 0 Then
                        Avanço(0) = 80
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 25) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 50) = "p"
                        GeraNotas2()
                    ElseIf NumeroAleatório = 73 AndAlso QtdePixeisCompasso(0) - 40 >= 0 Then
                        Avanço(0) = 40
                        ArrayX(MãoEsqDir, PosiçãoNoArray) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 10) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 20) = "x"
                        ArrayX(MãoEsqDir, PosiçãoNoArray + 30) = "xl"
                        GeraNotas2()
                    End If
                Loop
            Next

            'Array.Clear(PosiçãoNotas, 0, 63)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub GeraNotas2()

        Try

            Imagens(MãoEsqDir, PosiçãoNoArray) = DirectCast(myResourceManager.GetObject(c & "Nota" & NumeroAleatório), Image)
            ArrayXX(MãoEsqDir, PosiçãoNoArray) = NumeroAleatório
            PosiçãoNoArray += Avanço(0)
            QtdePixeisCompasso(0) -= Avanço(0)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub EstudoRitmico_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            WinMM.midiInReset(hMidiIn) 'se não resetar o midiinclose não funcionará
            WinMM.midiInClose(hMidiIn)
            MidiPlayer.CloseMidi()
        Catch ex As Exception
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

            microTimer.Enabled = False
        End Try

    End Sub

    Private Sub EstudoRitmico_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            MidiPlayer.OpenMidi()
        Catch ex As Exception
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

            AddHandler microTimer.MicroTimerElapsed, New MicroLibrary.MicroTimer.MicroTimerElapsedEventHandler(AddressOf OnTimedEvent)


            'Dim bmp As New Bitmap(1236, 739)
            'Dim graf As Graphics = Graphics.FromImage(bmp)
            'graf.DrawImage(My.Resources.Ajuda, 0, 0, My.Resources.Ajuda.Width, My.Resources.Ajuda.Height)
            'Me.BackgroundImage = bmp


            a = 0
            g = 0
            DistanciaInicialDoCompasso(1) = 0
            p(0) = -1
            'GerarExercícioRitmico()

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
        End Try

    End Sub

    Private Sub AtualizaRegiãoPercentualAcertos()
        Dim Rect31 As New Rectangle(590, 728, 155, 71)
        Me.Invalidate(Rect31)
    End Sub

    Private Sub OnTimedEvent(ByVal sender As Object, ByVal timerEventArgs As MicroLibrary.MicroTimerEventArgs)

        Try

            If a >= QtdeCompassos * QtdePixeisCompasso(1) Then

                If h = 0 Then
                    h = 100
                    If My.Settings.NovoValorMãoDireita = True Then MãoEsqDir = 0 : CalcularAcertosErros()
                    If My.Settings.NovoValorMãoEsquerda = True Then MãoEsqDir = 1 : CalcularAcertosErros()
                    AcertosErros(2) = CInt((AcertosErros(0) / (AcertosErros(0) + AcertosErros(1))) * 100)
                    AtualizaRegiãoDasNotas()
                    AtualizaRegiãoPercentualAcertos()
                    'MsgBox(AcertosErros(0) & " " & AcertosErros(1) & " " & AcertosErros(0) + AcertosErros(1))
                End If
                'microTimer.Enabled = False
            End If


            p(2) = a
            If ArrayX(4, a) = "x" Then
                p(1) = a
                If My.Settings.NovoValorLocalizaçãoSubdivisãoCompassoAtual = True Then AtualizaRegiãoRetangulo()
                DefinePosiçãoPonteiroMetronomo()
                AtualizaRegiãoMetrônomo()
            ElseIf ArrayX(4, a) = "xx" Then
                p(0) = a
                p(1) = a
                If My.Settings.NovoValorLocalizaçãoCompassoAtual = True Then AtualizaRegiãoRetangulo()
                DefinePosiçãoPonteiroMetronomo()
                AtualizaRegiãoMetrônomo()
                DistanciaInicialDoCompasso(1) += 20
                DistanciaInicialDoCompasso(2) += 20
            Else
                If My.Settings.NovoValorLocalizaçãoMicroSubdivisãoCompassoAtual = True Then AtualizaRegiãoRetangulo()
            End If



            If My.Settings.NovoValorTocarSonsMetronomo = True Then
                If h = 0 Then
                    If ArrayX(4, a) = "x" Then
                        Try
                            MidiPlayer.Play(New NoteOff(0, CType(My.Settings.NovoValorMetrônomoTempoFraco.Substring(0, 2), GeneralMidiPercussion), CByte(My.Settings.NovoValorVolumeTempoFraco)))
                            MidiPlayer.Play(New NoteOn(0, CType(My.Settings.NovoValorMetrônomoTempoFraco.Substring(0, 2), GeneralMidiPercussion), CByte(My.Settings.NovoValorVolumeTempoFraco)))
                        Catch ex As Exception

                            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
                        Finally
                            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

                        End Try
                    ElseIf ArrayX(4, a) = "xx" Then
                        Try
                            MidiPlayer.Play(New NoteOff(0, CType(My.Settings.NovoValorMetrônomoTempoForte.Substring(0, 2), GeneralMidiPercussion), CByte(My.Settings.NovoValorVolumeTempoForte)))
                            MidiPlayer.Play(New NoteOn(0, CType(My.Settings.NovoValorMetrônomoTempoForte.Substring(0, 2), GeneralMidiPercussion), CByte(My.Settings.NovoValorVolumeTempoForte)))
                        Catch ex As Exception

                            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
                        Finally
                            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

                        End Try
                    End If
                ElseIf h = 1 Then
                    If ArrayX(0, a) = "x" OrElse ArrayX(1, a) = "x" OrElse ArrayX(0, a) = "xl" OrElse ArrayX(1, a) = "xl" Then
                        If My.Settings.NovoValorTocarSonsMetronomo = True Then Som()
                    End If
                ElseIf h = 2 Then
                    If ArrayX(2, a) = "x" OrElse ArrayX(3, a) = "x" Then
                        If My.Settings.NovoValorTocarSonsMetronomo = True Then Som()
                    End If
                End If
            End If


            If a < QtdeCompassos * QtdePixeisCompasso(1) Then a += 5


            If p(2) = QtdePixeisCompasso(1) - 5 AndAlso l = 0 Then
                a = 0 'retorna ao inicio do primeiro compasso, até que alguma tecla seja tocada
                p(0) = -1
                p(1) = -1
                p(2) = -1
                DistanciaInicialDoCompasso(1) = 0
                DistanciaInicialDoCompasso(2) = 0
            End If


            'MsgBox(String.Format("Count = {0:#,0}  Timer = {1:#,0} µs, LateBy = {2:#,0} µs, ExecutionTime = {3:#,0} µs", timerEventArgs.TimerCount, timerEventArgs.ElapsedMicroseconds, timerEventArgs.TimerLateBy, timerEventArgs.CallbackFunctionExecutionTime))

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub AtualizaRegiãoRetangulo()

        Try

            If a >= 0 AndAlso a < LimiteCompasso Then

                Dim Rect As New Rectangle(70, 199, 1140, 12)
                Me.Invalidate(Rect)

            ElseIf a >= LimiteCompasso AndAlso a < LimiteCompasso * 2 Then

                Dim Rect As New Rectangle(70, 199, 1140, 12)
                Me.Invalidate(Rect)

                Dim Rect2 As New Rectangle(70, 344, 1140, 12)
                Me.Invalidate(Rect2)

            ElseIf a >= LimiteCompasso * 2 AndAlso a < LimiteCompasso * 3 Then

                Dim Rect2 As New Rectangle(70, 344, 1140, 12)
                Me.Invalidate(Rect2)

                Dim Rect3 As New Rectangle(70, 489, 1140, 12)
                Me.Invalidate(Rect3)

            ElseIf a >= LimiteCompasso * 3 AndAlso a < LimiteCompasso * 4 Then

                Dim Rect3 As New Rectangle(70, 489, 1140, 12)
                Me.Invalidate(Rect3)

                Dim Rect4 As New Rectangle(70, 634, 1140, 12)
                Me.Invalidate(Rect4)

            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub CalcularAcertosErros()

        Try

            For i = 0 To QtdeCompassos * QtdePixeisCompasso(1)
                j = i
                If i - 5 < 0 Then Valor(1) = 0 Else Valor(1) = i - 5
                If i - 10 < 0 Then Valor(2) = 0 Else Valor(2) = i - 10

                If ArrayX(MãoEsqDir, i) = "" OrElse ArrayX(MãoEsqDir, i) = "p" Then
                    If ArrayX(MãoEsqDir + 2, i) = "x" AndAlso (ArrayX(MãoEsqDir, Valor(1)) = "" OrElse ArrayX(MãoEsqDir, Valor(1)) = "p") AndAlso (ArrayX(MãoEsqDir, Valor(2)) = "" OrElse ArrayX(MãoEsqDir, Valor(2)) = "p") AndAlso (ArrayX(MãoEsqDir, i + 5) = "" OrElse ArrayX(MãoEsqDir, i + 5) = "p") AndAlso (ArrayX(MãoEsqDir, i + 10) = "" OrElse ArrayX(MãoEsqDir, i + 10) = "p") Then
                        ArrayX(MãoEsqDir + 5, i) = "e"
                        AcertosErros(1) += 1
                    End If
                    If ArrayX(MãoEsqDir, i) = "p" AndAlso ArrayX(MãoEsqDir + 2, i) = "" Then
                        ArrayX(MãoEsqDir + 5, i) = "c"
                        AcertosErros(0) += 1
                        jj = i
                        CorrigeAcertosErrosCasoExistaDinâmica()
                    End If
                Else
                    If ArrayX(MãoEsqDir + 2, i) = "x" Then
                        ArrayX(MãoEsqDir + 5, i) = "c"
                        AcertosErros(0) += 1
                        jj = i
                        CorrigeAcertosErrosCasoExistaDinâmica()
                    ElseIf ArrayX(MãoEsqDir + 2, Valor(1)) = "x" Then
                        ArrayX(MãoEsqDir + 5, Valor(1)) = "c"
                        AcertosErros(0) += 1
                        If My.Settings.NovoValorSevera = True Then ArrayX(MãoEsqDir + 5, Valor(1)) = "e" : AcertosErros(1) += 1 : AcertosErros(0) -= 1
                        jj = Valor(1)
                        CorrigeAcertosErrosCasoExistaDinâmica()
                    ElseIf ArrayX(MãoEsqDir + 2, i + 5) = "x" Then
                        ArrayX(MãoEsqDir + 5, i + 5) = "c"
                        AcertosErros(0) += 1
                        If My.Settings.NovoValorSevera = True Then ArrayX(MãoEsqDir + 5, i + 5) = "e" : AcertosErros(1) += 1 : AcertosErros(0) -= 1
                        jj = i + 5
                        CorrigeAcertosErrosCasoExistaDinâmica()
                    ElseIf ArrayX(MãoEsqDir + 2, Valor(2)) = "x" Then
                        ArrayX(MãoEsqDir + 5, Valor(2)) = "c"
                        AcertosErros(0) += 1
                        If My.Settings.NovoValorSevera = True OrElse My.Settings.NovoValorNormal = True Then ArrayX(MãoEsqDir + 5, Valor(2)) = "e" : AcertosErros(1) += 1 : AcertosErros(0) -= 1
                        jj = Valor(2)
                        CorrigeAcertosErrosCasoExistaDinâmica()
                    ElseIf ArrayX(MãoEsqDir + 2, i + 10) = "x" Then
                        ArrayX(MãoEsqDir + 5, i + 10) = "c"
                        AcertosErros(0) += 1
                        If My.Settings.NovoValorSevera = True OrElse My.Settings.NovoValorNormal = True Then ArrayX(MãoEsqDir + 5, i + 10) = "e" : AcertosErros(1) += 1 : AcertosErros(0) -= 1
                        jj = i + 10
                        CorrigeAcertosErrosCasoExistaDinâmica()
                    Else
                        ArrayX(MãoEsqDir + 5, i) = "e"
                        AcertosErros(1) += 1
                    End If
                End If

                i += 4
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub CorrigeAcertosErrosCasoExistaDinâmica()

        Try

            If My.Settings.NovoValorDinamicaMD = True OrElse My.Settings.NovoValorDinamicaME = True Then
                If ArrayX(MãoEsqDir + 15, jj) <> ArrayX(MãoEsqDir + 17, j) Then
                    ArrayX(MãoEsqDir + 5, jj) = "ed"
                    AcertosErros(1) += 1
                    AcertosErros(0) -= 1 'como houve erro e não acerto, é preciso descontar o acerto que havia sido considerado
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub AtualizaRegiãoMetrônomo()
        Dim Rect21 As New Rectangle(1068, 58, 40, 40)
        Me.Invalidate(Rect21)
    End Sub

    Private Sub Som()

        Try

            MidiPlayer.Play(New NoteOff(0, CType(My.Settings.NovoValorSomRitmo.Substring(0, 2), GeneralMidiPercussion), CByte(My.Settings.NovoValorVolumeRitmo)))
            MidiPlayer.Play(New NoteOn(0, CType(My.Settings.NovoValorSomRitmo.Substring(0, 2), GeneralMidiPercussion), CByte(My.Settings.NovoValorVolumeRitmo)))
            'Thread(0) = New Thread(AddressOf SegundaThreadCode)
            'Thread(0).Name = "Segunda Thread"
            'Thread(0).Start()

        Catch ex As Exception

            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub SegundaThreadCode()
        Try
            MidiPlayer.Play(New NoteOn(0, GeneralMidiPercussion.HandClap, 127))
        Catch ex As Exception

            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Som2()
        Thread(1) = New Thread(AddressOf TerceiraThreadCode)
        Thread(1).Name = "Terceira Thread"
        Thread(1).Start()
    End Sub

    Private Sub TerceiraThreadCode()
        Try
            MidiPlayer.Play(New NoteOn(0, ValorSom, CByte(ValorVolumeSom)))
        Catch ex As Exception

            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub MemoNotes_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown, Button2.KeyDown, Button1.KeyDown, Button3.KeyDown, Button4.KeyDown, NumericUpDown1.KeyDown

        Try

            ' A tecla pressionada está na nossa lista? 
            If Not lista.Contains(e.KeyCode) Then
                ' Não! A adicionamos à lista.
                lista.Add(e.KeyCode)
                If Chr(e.KeyCode) = My.Settings.NovoValorTeclaMãoDireita AndAlso My.Settings.NovoValorMãoDireita Then
                    If l = 0 Then IniciarJogo()
                    ArrayX(2, p(2)) = "x"
                    'Desenhar()
                    If My.Settings.NovoValorTocarSonsMetronomo = True Then Som()
                    AtualizaRegiãoDasMarcações1()
                End If
                If Chr(e.KeyCode) = My.Settings.NovoValorTeclaMãoEsquerda AndAlso My.Settings.NovoValorMãoEsquerda Then
                    If l = 0 Then IniciarJogo()
                    ArrayX(3, p(2)) = "x"
                    ' Desenhar()
                    If My.Settings.NovoValorTocarSonsMetronomo = True Then Som()
                    AtualizaRegiãoDasMarcações2()
                End If
            End If

        Catch ex As Exception
            If Chr(e.KeyCode) = My.Settings.NovoValorTeclaMãoDireita AndAlso My.Settings.NovoValorMãoDireita Then ArrayX(2, 0) = "x" : Me.Refresh()
            If Chr(e.KeyCode) = My.Settings.NovoValorTeclaMãoEsquerda AndAlso My.Settings.NovoValorMãoEsquerda Then ArrayX(3, 0) = "x" : Me.Refresh()
            'MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui

        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub IniciarJogo()
        l = 1
    End Sub

    Private Sub MemoNotes_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyUp, Button2.KeyUp, Button1.KeyUp, Button3.KeyUp, Button4.KeyUp, NumericUpDown1.KeyUp

        Try

            'If e.KeyCode = Keys.A Then
            'lista.Remove(Keys.A)
            'ElseIf e.KeyCode = Keys.L Then
            'lista.Remove(Keys.L)
            'End If
            lista.Remove(e.KeyCode)
            MidiPlayer.Play(New NoteOff(0, CType(My.Settings.NovoValorSomRitmo.Substring(0, 2), GeneralMidiPercussion), CByte(My.Settings.NovoValorVolumeRitmo)))

        Catch ex As Exception

            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub EstudoRitmico_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click, Me.Load
        Try

            microTimer.Enabled = False
            GerarExercícioRitmico()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub ab_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown, Label3.MouseDown, Label2.MouseDown, Label1.MouseDown
        VGa = Me.MousePosition.X - Me.Location.X
        VGb = Me.MousePosition.Y - Me.Location.Y
    End Sub

    Private Sub ab_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove, Label3.MouseMove, Label2.MouseMove, Label1.MouseMove

        If e.Button = MouseButtons.Left Then
            VGnewPoint = Me.MousePosition
            VGnewPoint.X -= VGa
            VGnewPoint.Y -= VGb
            Me.Location = VGnewPoint
        End If

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
                If [KeyVelocity] >= 0 AndAlso [KeyVelocity] <= 15 Then
                    Din = "ppp"
                ElseIf [KeyVelocity] >= 16 AndAlso [KeyVelocity] <= 31 Then
                    Din = "pp"
                ElseIf [KeyVelocity] >= 32 AndAlso [KeyVelocity] <= 47 Then
                    Din = "p"
                ElseIf [KeyVelocity] >= 48 AndAlso [KeyVelocity] <= 63 Then
                    Din = "mp"
                ElseIf [KeyVelocity] >= 64 AndAlso [KeyVelocity] <= 79 Then
                    Din = "mf"
                ElseIf [KeyVelocity] >= 80 AndAlso [KeyVelocity] <= 95 Then
                    Din = "f"
                ElseIf [KeyVelocity] >= 96 AndAlso [KeyVelocity] <= 111 Then
                    Din = "ff"
                ElseIf [KeyVelocity] >= 112 AndAlso [KeyVelocity] <= 127 Then
                    Din = "fff"
                End If

                If [Param] - 20 >= 40 AndAlso My.Settings.NovoValorMãoDireita Then
                    If l = 0 Then IniciarJogo()
                    ArrayX(2, p(2)) = "x"
                    ArrayX(15, p(2)) = Din
                    'Desenhar()
                    If My.Settings.NovoValorTocarSonsMetronomo = True Then Som()
                    AtualizaRegiãoDasMarcações1()
                End If
                If [Param] - 20 <= 39 AndAlso My.Settings.NovoValorMãoEsquerda Then
                    If l = 0 Then IniciarJogo()
                    ArrayX(3, p(2)) = "x"
                    ArrayX(16, p(2)) = Din
                    ' Desenhar()
                    If My.Settings.NovoValorTocarSonsMetronomo = True Then Som()
                    AtualizaRegiãoDasMarcações2()
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub TouchOff(ByVal [Param] As Byte, ByVal [KeyVelocity] As Byte)

        Try
            'If Me.KeyCol([Param]).InvokeRequired Then
            'Me.KeyCol([Param]).Invoke(DelgParamOff, New Object() {[Param], [KeyVelocity]})
            ' Else
            'Me.KeyCol([Param]).Visible = False

            'End If
            MidiPlayer.Play(New NoteOff(0, CType(My.Settings.NovoValorSomRitmo.Substring(0, 2), GeneralMidiPercussion), CByte(My.Settings.NovoValorVolumeRitmo)))
        Catch ex As Exception

            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub AtualizaRegiãoDasMarcações1()

        Try

            CálculoRegiãoMarcações()

            Dim Rect As New Rectangle(a - CompensaçãoLeft(2) + 37 + DistanciaInicialDoCompasso(2), 184 + CompensaçãoTop(2), 70, 8)
            Me.Invalidate(Rect)
            'MsgBox(a - CompensaçãoLeft(2) + 107 + DistanciaInicialDoCompasso(2) & " " & 184 + CompensaçãoTop(2) & " " & " 10 8")

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub AtualizaRegiãoDasMarcações2()

        Try

            CálculoRegiãoMarcações()

            Dim Rect As New Rectangle(a - CompensaçãoLeft(2) + 37 + DistanciaInicialDoCompasso(2), 221 + CompensaçãoTop(2), 70, 8)
            Me.Invalidate(Rect)
            'MsgBox(a - CompensaçãoLeft(2) + 107 + DistanciaInicialDoCompasso(2) & " " & 184 + CompensaçãoTop(2) & " " & " 10 8")

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub CálculoRegiãoMarcações()

        Try

            If a >= 0 AndAlso a < LimiteCompasso Then
                CompensaçãoLeft(2) = 0
                CompensaçãoTop(2) = 0
            ElseIf a >= LimiteCompasso AndAlso a < LimiteCompasso * 2 Then
                CompensaçãoLeft(2) = CInt(LimiteCompasso + ((QtdeCompassos / 4) * 20))
                CompensaçãoTop(2) = 145
            ElseIf a >= LimiteCompasso * 2 AndAlso a < LimiteCompasso * 3 Then
                CompensaçãoLeft(2) = CInt((LimiteCompasso * 2) + ((QtdeCompassos / 4) * 40))
                CompensaçãoTop(2) = 290
            ElseIf a >= LimiteCompasso * 3 AndAlso a < LimiteCompasso * 4 Then
                CompensaçãoLeft(2) = CInt((LimiteCompasso * 3) + ((QtdeCompassos / 4) * 60))
                CompensaçãoTop(2) = 435
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

            a = 0
            h = 1
            DistanciaInicialDoCompasso(1) = 0
            microTimer.Interval = CInt((60000000 / BPM / QtdeIntervalosSubdivisão) * AjusteBPM)
            microTimer.Enabled = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub NumericUpDown1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown1.ValueChanged
        AtualizaRegiãoDoTempo()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Try

            a = 0
            h = 2
            DistanciaInicialDoCompasso(1) = 0
            microTimer.Interval = CInt((60000000 / BPM / QtdeIntervalosSubdivisão) * AjusteBPM)
            microTimer.Enabled = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub DefinePosiçãoPonteiroMetronomo()

        Try

            If g = 0 Then
                g = 1
            Else
                g = 0
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub Label5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label5.Click
        Me.WindowState = CType(1, FormWindowState) '1 é para minimizar
    End Sub

    Private Sub Label4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label4.Click
        Try

            Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Try

            microTimer.Enabled = False
            FigurasRitmicas.ShowDialog()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub
End Class