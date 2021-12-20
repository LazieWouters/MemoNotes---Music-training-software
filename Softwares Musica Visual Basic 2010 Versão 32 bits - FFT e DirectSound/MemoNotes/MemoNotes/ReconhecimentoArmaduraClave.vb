Public Class ReconhecimentoArmaduraClave
    Inherits PerPixelAlphaForm 'para mecher no form no modo design é preciso mudar para 'Inherits Form'
    'e depois de ajustado voltar para 'Inherits PerPixelAlphaForm' e clicar para mudar o inherit
    Public TransAmount As Byte = 255


    Dim ArmaduraClave, ValorAleatórioEscolhido, ValorCorreto, ValorIncorreto, ValorTotal, ValorFinal, ValorInicial, LeftButton, TopButton, PosicaoX6, PosicaoY6, WidthButton As Integer
    Dim CorretoIncorreto, Tonalidade(2) As String

    Private Sub GerarAramduraDeClave()

        Try

            ArmaduraClave = 15


            Dim gr As Graphics
            Dim FaceBit As New Bitmap(My.Resources.ReconhecimentoArmaduraClave)
            gr = Graphics.FromImage(FaceBit)
            gr.SmoothingMode = SmoothingMode.AntiAlias


            'desenha tick
            If CheckBox1.Checked Then
                gr.DrawImage(My.Resources.Tick, 683, 141, 45, 43)
            Else
                gr.DrawImage(My.Resources.Tick, 683, 176, 45, 43)
                gr.DrawImage(My.Resources.ArmadurasTelaReconhecimento, 128, 296, 811, 228)
            End If


            Dim Branco As SolidBrush = New SolidBrush(Color.FromArgb(255, 255, 250))
            Dim Preto As SolidBrush = New SolidBrush(Color.FromArgb(0, 0, 0))
            Dim Laranja As SolidBrush = New SolidBrush(Color.FromArgb(255, 153, 0))
            Dim Azul As SolidBrush = New SolidBrush(Color.FromArgb(0, 153, 255))
            Dim Fonte As New Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel)
            Dim Fonte2 As New Font("Arial", 15, FontStyle.Regular, GraphicsUnit.Pixel)
            Dim Fonte3 As New Font("Arial", 10, FontStyle.Regular, GraphicsUnit.Pixel)
            Dim Fonte4 As New Font("Arial", 30, FontStyle.Bold, GraphicsUnit.Pixel)
            Dim MedidasTexto As SizeF = gr.MeasureString(CorretoIncorreto, Fonte)
            Dim MedidasTexto2 As SizeF = gr.MeasureString("Total de Acertos: " & ValorFinal & "%", Fonte2)
            Dim MedidasTexto3 As SizeF = gr.MeasureString("Acertos: " & ValorCorreto & "   Erros: " & ValorIncorreto & "   Total: " & ValorTotal, Fonte3)
            Dim MedidasTexto6 As SizeF = gr.MeasureString(Tonalidade(2), Fonte3)
            Dim PosicaoX As Integer = CInt((Me.Width / 2) - (MedidasTexto.Width / 2))
            Dim PosicaoY As Integer = 180
            Dim PosicaoX2 As Integer = CInt((Me.Width / 2) - (MedidasTexto2.Width / 2))
            Dim PosicaoY2 As Integer = 220
            Dim PosicaoX3 As Integer = CInt(((Me.Width - 8) / 2) - (MedidasTexto3.Width / 2))
            Dim PosicaoY3 As Integer = 235
            PosicaoX6 = CInt((LeftButton + (WidthButton / 2)) - (MedidasTexto6.Width / 2))
            PosicaoY6 = TopButton + 65

            If ValorInicial <> 0 Then
                gr.DrawString(CorretoIncorreto, Fonte, Azul, PosicaoX, PosicaoY)
                gr.DrawString("Total de Acertos: " & ValorFinal & "%", Fonte2, Azul, PosicaoX2, PosicaoY2)
                gr.DrawString("Acertos: " & ValorCorreto & "   Erros: " & ValorIncorreto & "   Total: " & ValorTotal, Fonte3, Azul, PosicaoX3, PosicaoY3)

                If CorretoIncorreto = "Correto" Then
                    gr.DrawImage(My.Resources.CorretoArmaduraClave, LeftButton, TopButton - 14, 126, 81)
                    gr.DrawString(Tonalidade(2), Fonte3, Preto, PosicaoX6, PosicaoY6)
                ElseIf CorretoIncorreto = "Incorreto" Then
                    gr.DrawImage(My.Resources.IncorretoArmaduraClave, LeftButton, TopButton - 14, 126, 81)
                    If ValorAleatórioEscolhido = 0 Then
                        gr.DrawImage(My.Resources.CorretoArmaduraClave, Button1.Left, Button1.Top - 14, 126, 81)
                        gr.DrawString(Tonalidade(2), Fonte3, Preto, CInt((Button1.Left + (Button1.Width / 2)) - (MedidasTexto6.Width / 2)), Button1.Top + 65)
                    ElseIf ValorAleatórioEscolhido = 1 Then
                        gr.DrawImage(My.Resources.CorretoArmaduraClave, Button2.Left, Button2.Top - 14, 126, 81)
                        gr.DrawString(Tonalidade(2), Fonte3, Preto, CInt((Button2.Left + (Button2.Width / 2)) - (MedidasTexto6.Width / 2)), Button2.Top + 65)
                    ElseIf ValorAleatórioEscolhido = 2 Then
                        gr.DrawImage(My.Resources.CorretoArmaduraClave, Button3.Left, Button3.Top - 14, 126, 81)
                        gr.DrawString(Tonalidade(2), Fonte3, Preto, CInt((Button3.Left + (Button3.Width / 2)) - (MedidasTexto6.Width / 2)), Button3.Top + 65)
                    ElseIf ValorAleatórioEscolhido = 3 Then
                        gr.DrawImage(My.Resources.CorretoArmaduraClave, Button4.Left, Button4.Top - 14, 126, 81)
                        gr.DrawString(Tonalidade(2), Fonte3, Preto, CInt((Button4.Left + (Button4.Width / 2)) - (MedidasTexto6.Width / 2)), Button4.Top + 65)
                    ElseIf ValorAleatórioEscolhido = 4 Then
                        gr.DrawImage(My.Resources.CorretoArmaduraClave, Button5.Left, Button5.Top - 14, 126, 81)
                        gr.DrawString(Tonalidade(2), Fonte3, Preto, CInt((Button5.Left + (Button5.Width / 2)) - (MedidasTexto6.Width / 2)), Button5.Top + 65)
                    ElseIf ValorAleatórioEscolhido = 5 Then
                        gr.DrawImage(My.Resources.CorretoArmaduraClave, Button10.Left, Button10.Top - 14, 126, 81)
                        gr.DrawString(Tonalidade(2), Fonte3, Preto, CInt((Button10.Left + (Button10.Width / 2)) - (MedidasTexto6.Width / 2)), Button10.Top + 65)
                    ElseIf ValorAleatórioEscolhido = 6 Then
                        gr.DrawImage(My.Resources.CorretoArmaduraClave, Button9.Left, Button9.Top - 14, 126, 81)
                        gr.DrawString(Tonalidade(2), Fonte3, Preto, CInt((Button9.Left + (Button9.Width / 2)) - (MedidasTexto6.Width / 2)), Button9.Top + 65)
                    ElseIf ValorAleatórioEscolhido = 7 Then
                        gr.DrawImage(My.Resources.CorretoArmaduraClave, Button8.Left, Button8.Top - 14, 126, 81)
                        gr.DrawString(Tonalidade(2), Fonte3, Preto, CInt((Button8.Left + (Button8.Width / 2)) - (MedidasTexto6.Width / 2)), Button8.Top + 65)
                    ElseIf ValorAleatórioEscolhido = 8 Then
                        gr.DrawImage(My.Resources.CorretoArmaduraClave, Button7.Left, Button7.Top - 14, 126, 81)
                        gr.DrawString(Tonalidade(2), Fonte3, Preto, CInt((Button7.Left + (Button7.Width / 2)) - (MedidasTexto6.Width / 2)), Button7.Top + 65)
                    ElseIf ValorAleatórioEscolhido = 9 Then
                        gr.DrawImage(My.Resources.CorretoArmaduraClave, Button6.Left, Button6.Top - 14, 126, 81)
                        gr.DrawString(Tonalidade(2), Fonte3, Preto, CInt((Button6.Left + (Button6.Width / 2)) - (MedidasTexto6.Width / 2)), Button6.Top + 65)
                    ElseIf ValorAleatórioEscolhido = 10 Then
                        gr.DrawImage(My.Resources.CorretoArmaduraClave, Button15.Left, Button15.Top - 14, 126, 81)
                        gr.DrawString(Tonalidade(2), Fonte3, Preto, CInt((Button15.Left + (Button15.Width / 2)) - (MedidasTexto6.Width / 2)), Button15.Top + 65)
                    ElseIf ValorAleatórioEscolhido = 11 Then
                        gr.DrawImage(My.Resources.CorretoArmaduraClave, Button14.Left, Button14.Top - 14, 126, 81)
                        gr.DrawString(Tonalidade(2), Fonte3, Preto, CInt((Button14.Left + (Button14.Width / 2)) - (MedidasTexto6.Width / 2)), Button14.Top + 65)
                    ElseIf ValorAleatórioEscolhido = 12 Then
                        gr.DrawImage(My.Resources.CorretoArmaduraClave, Button13.Left, Button13.Top - 14, 126, 81)
                        gr.DrawString(Tonalidade(2), Fonte3, Preto, CInt((Button13.Left + (Button13.Width / 2)) - (MedidasTexto6.Width / 2)), Button13.Top + 65)
                    ElseIf ValorAleatórioEscolhido = 13 Then
                        gr.DrawImage(My.Resources.CorretoArmaduraClave, Button12.Left, Button12.Top - 14, 126, 81)
                        gr.DrawString(Tonalidade(2), Fonte3, Preto, CInt((Button12.Left + (Button12.Width / 2)) - (MedidasTexto6.Width / 2)), Button12.Top + 65)
                    ElseIf ValorAleatórioEscolhido = 14 Then
                        gr.DrawImage(My.Resources.CorretoArmaduraClave, Button11.Left, Button11.Top - 14, 126, 81)
                        gr.DrawString(Tonalidade(2), Fonte3, Preto, CInt((Button11.Left + (Button11.Width / 2)) - (MedidasTexto6.Width / 2)), Button11.Top + 65)
                    End If
                End If

            End If

            Do While ArmaduraClave > 14 OrElse ArmaduraClave = ValorAleatórioEscolhido

                ' Create a byte array to hold the random value.
                Dim randomNumber(0) As Byte

                ' Create a new instance of the RNGCryptoServiceProvider.
                Dim Gen As New Security.Cryptography.RNGCryptoServiceProvider()

                ' Fill the array with a random value.
                Gen.GetBytes(randomNumber)

                ' Convert the byte to an integer value to make the modulus operation easier.
                ArmaduraClave = Int(Convert.ToInt32(randomNumber(0)))

            Loop

            ValorAleatórioEscolhido = ArmaduraClave

            If CheckBox1.Checked Then
                gr.DrawImage(My.Resources.ClaveC, 489, 130, 88, 52)
                If ArmaduraClave = 0 Then 'Clave de Dó
                    Tonalidade(2) = "C-D-E-F-G-A-B-C"
                ElseIf ArmaduraClave = 1 Then 'Clave de Dó#
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 489 + 30, 130 + 10, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 489 + 37, 130 + 18, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 489 + 44, 130 + 7, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 489 + 51, 130 + 15, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 489 + 58, 130 + 23, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 489 + 65, 130 + 13, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 489 + 72, 130 + 20, 6, 16)
                    Tonalidade(2) = "C#-D#-E#-F#-G#-A#-B#-C#"
                ElseIf ArmaduraClave = 2 Then 'Clave de Réb
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 489 + 30, 130 + 17, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 489 + 37, 130 + 10, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 489 + 44, 130 + 20, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 489 + 51, 130 + 12, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 489 + 58, 130 + 22, 6, 16)
                    Tonalidade(2) = "Db-Eb-F-Gb-Ab-Bb-C-Db"
                ElseIf ArmaduraClave = 3 Then 'Clave de Ré
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 489 + 30, 130 + 10, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 489 + 37, 130 + 18, 6, 16)
                    Tonalidade(2) = "D-E-F#-G-A-B-C#-D"
                ElseIf ArmaduraClave = 4 Then 'Clave de Mib
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 489 + 30, 130 + 17, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 489 + 37, 130 + 10, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 489 + 44, 130 + 20, 6, 16)
                    Tonalidade(2) = "Eb-F-G-Ab-Bb-C-D-Eb"
                ElseIf ArmaduraClave = 5 Then 'Clave de Mi
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 489 + 30, 130 + 10, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 489 + 37, 130 + 18, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 489 + 44, 130 + 7, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 489 + 51, 130 + 15, 6, 16)
                    Tonalidade(2) = "E-F#-G#-A-B-C#-D#-E"
                ElseIf ArmaduraClave = 6 Then 'Clave de Fá
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 489 + 30, 130 + 17, 6, 16)
                    Tonalidade(2) = "F-G-A-Bb-C-D-E-F"
                ElseIf ArmaduraClave = 7 Then 'Clave de Fá#
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 489 + 30, 130 + 10, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 489 + 37, 130 + 18, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 489 + 44, 130 + 7, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 489 + 51, 130 + 15, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 489 + 58, 130 + 23, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 489 + 65, 130 + 13, 6, 16)
                    Tonalidade(2) = "F#-G#-A#-B-C#-D#-E#-F#"
                ElseIf ArmaduraClave = 8 Then 'Clave de Solb
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 489 + 30, 130 + 17, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 489 + 37, 130 + 10, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 489 + 44, 130 + 20, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 489 + 51, 130 + 12, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 489 + 58, 130 + 22, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 489 + 65, 130 + 15, 6, 16)
                    Tonalidade(2) = "Gb-Ab-Bb-Cb-Db-Eb-F-Gb"
                ElseIf ArmaduraClave = 9 Then 'Clave de Sol
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 489 + 30, 130 + 10, 6, 16)
                    Tonalidade(2) = "G-A-B-C-D-E-F#-G"
                ElseIf ArmaduraClave = 10 Then 'Clave de Láb
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 489 + 30, 130 + 17, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 489 + 37, 130 + 10, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 489 + 44, 130 + 20, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 489 + 51, 130 + 12, 6, 16)
                    Tonalidade(2) = "Ab-Bb-C-Db-Eb-F-G-Ab"
                ElseIf ArmaduraClave = 11 Then 'Clave de Lá
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 489 + 30, 130 + 10, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 489 + 37, 130 + 18, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 489 + 44, 130 + 7, 6, 16)
                    Tonalidade(2) = "A-B-C#-D-E-F#-G#-A"
                ElseIf ArmaduraClave = 12 Then 'Clave de Sib
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 489 + 30, 130 + 17, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 489 + 37, 130 + 10, 6, 16)
                    Tonalidade(2) = "Bb-C-D-Eb-F-G-A-Bb"
                ElseIf ArmaduraClave = 13 Then 'Clave de Si
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 489 + 30, 130 + 10, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 489 + 37, 130 + 18, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 489 + 44, 130 + 7, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 489 + 51, 130 + 15, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 489 + 58, 130 + 23, 6, 16)
                    Tonalidade(2) = "B-C#-D#-E-F#-G#-A#-B"
                ElseIf ArmaduraClave = 14 Then 'Clave de Dób
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 489 + 30, 130 + 17, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 489 + 37, 130 + 10, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 489 + 44, 130 + 20, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 489 + 51, 130 + 12, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 489 + 58, 130 + 22, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 489 + 65, 130 + 15, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 489 + 72, 130 + 25, 6, 16)
                    Tonalidade(2) = "Cb-Db-Eb-Fb-Gb-Ab-Bb-Cb"
                End If
            Else
                If ArmaduraClave = 0 Then '0#
                    Tonalidade(0) = "C" : Tonalidade(1) = "Am" : Tonalidade(2) = "C-D-E-F-G-A-B-C"
                ElseIf ArmaduraClave = 1 Then '1#
                    Tonalidade(0) = "G" : Tonalidade(1) = "Em" : Tonalidade(2) = "G-A-B-C-D-E-F#-G"
                ElseIf ArmaduraClave = 2 Then '2#
                    Tonalidade(0) = "D" : Tonalidade(1) = "Bm" : Tonalidade(2) = "D-E-F#-G-A-B-C#-D"
                ElseIf ArmaduraClave = 3 Then '3#
                    Tonalidade(0) = "A" : Tonalidade(1) = "F#m" : Tonalidade(2) = "A-B-C#-D-E-F#-G#-A"
                ElseIf ArmaduraClave = 4 Then '4#
                    Tonalidade(0) = "E" : Tonalidade(1) = "C#m" : Tonalidade(2) = "E-F#-G#-A-B-C#-D#-E"
                ElseIf ArmaduraClave = 5 Then '5#
                    Tonalidade(0) = "B" : Tonalidade(1) = "G#m" : Tonalidade(2) = "B-C#-D#-E-F#-G#-A#-B"
                ElseIf ArmaduraClave = 6 Then '6#
                    Tonalidade(0) = "F#" : Tonalidade(1) = "D#m" : Tonalidade(2) = "F#-G#-A#-B-C#-D#-E#-F#"
                ElseIf ArmaduraClave = 7 Then '7#
                    Tonalidade(0) = "C#" : Tonalidade(1) = "A#m" : Tonalidade(2) = "C#-D#-E#-F#-G#-A#-B#-C#"
                ElseIf ArmaduraClave = 8 Then '1b
                    Tonalidade(0) = "F" : Tonalidade(1) = "Dm" : Tonalidade(2) = "F-G-A-Bb-C-D-E-F"
                ElseIf ArmaduraClave = 9 Then '2b
                    Tonalidade(0) = "Bb" : Tonalidade(1) = "Gm" : Tonalidade(2) = "Bb-C-D-Eb-F-G-A-Bb"
                ElseIf ArmaduraClave = 10 Then '3b
                    Tonalidade(0) = "Eb" : Tonalidade(1) = "Cm" : Tonalidade(2) = "Eb-F-G-Ab-Bb-C-D-Eb"
                ElseIf ArmaduraClave = 11 Then '4b
                    Tonalidade(0) = "Ab" : Tonalidade(1) = "Fm" : Tonalidade(2) = "Ab-Bb-C-Db-Eb-F-G-Ab"
                ElseIf ArmaduraClave = 12 Then '5b
                    Tonalidade(0) = "Db" : Tonalidade(1) = "Bbm" : Tonalidade(2) = "Db-Eb-F-Gb-Ab-Bb-C-Db"
                ElseIf ArmaduraClave = 13 Then '6b
                    Tonalidade(0) = "Gb" : Tonalidade(1) = "Ebm" : Tonalidade(2) = "Gb-Ab-Bb-Cb-Db-Eb-F-Gb"
                ElseIf ArmaduraClave = 14 Then '7b
                    Tonalidade(0) = "Cb" : Tonalidade(1) = "Abm" : Tonalidade(2) = "Cb-Db-Eb-Fb-Gb-Ab-Bb-Cb"
                End If

                Dim MedidasTexto4 As SizeF = gr.MeasureString(Tonalidade(0), Fonte4)
                Dim MedidasTexto5 As SizeF = gr.MeasureString(Tonalidade(1), Fonte4)
                Dim PosicaoX4 As Integer = CInt((Me.Width / 2) - (MedidasTexto4.Width / 2))
                Dim PosicaoY4 As Integer = 115
                Dim PosicaoX5 As Integer = CInt((Me.Width / 2) - (MedidasTexto5.Width / 2))
                Dim PosicaoY5 As Integer = 150
                gr.DrawString(Tonalidade(0), Fonte4, Laranja, PosicaoX4, PosicaoY4)
                gr.DrawString(Tonalidade(1), Fonte4, Branco, PosicaoX5, PosicaoY5)
            End If


            Me.SetBitmap(FaceBit, TransAmount)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub ab_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown
        VGa = Me.MousePosition.X - Me.Location.X
        VGb = Me.MousePosition.Y - Me.Location.Y
    End Sub

    Private Sub ab_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove

        If e.Button = MouseButtons.Left Then
            VGnewPoint = Me.MousePosition
            VGnewPoint.X -= VGa
            VGnewPoint.Y -= VGb
            Me.Location = VGnewPoint
        End If

    End Sub

    Private Sub ReconhecimentoArmaduraClave_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ValorInicial = 0
        GerarAramduraDeClave()
    End Sub

    Private Sub IdentificarAcertosErros(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Button1.Click, Button2.Click, Button3.Click, Button4.Click, Button5.Click, Button6.Click, Button7.Click, Button8.Click, Button9.Click, Button10.Click, _
        Button11.Click, Button12.Click, Button13.Click, Button14.Click, Button15.Click

        Try

            If ValorInicial = 0 Then ValorInicial = 1

            ' Obtém referência ao botão que invocou este método.
            Dim pbutton As Button = DirectCast(sender, Button)

            LeftButton = pbutton.Left
            TopButton = pbutton.Top
            WidthButton = pbutton.Width


            If (pbutton.Text = "0" AndAlso ArmaduraClave = 0) OrElse (pbutton.Text = "1" AndAlso ArmaduraClave = 1) OrElse (pbutton.Text = "2" AndAlso ArmaduraClave = 2) OrElse (pbutton.Text = "3" AndAlso ArmaduraClave = 3) _
            OrElse (pbutton.Text = "4" AndAlso ArmaduraClave = 4) OrElse (pbutton.Text = "5" AndAlso ArmaduraClave = 5) OrElse (pbutton.Text = "6" AndAlso ArmaduraClave = 6) OrElse (pbutton.Text = "7" AndAlso ArmaduraClave = 7) _
            OrElse (pbutton.Text = "8" AndAlso ArmaduraClave = 8) OrElse (pbutton.Text = "9" AndAlso ArmaduraClave = 9) OrElse (pbutton.Text = "10" AndAlso ArmaduraClave = 10) OrElse (pbutton.Text = "11" AndAlso ArmaduraClave = 11) _
            OrElse (pbutton.Text = "12" AndAlso ArmaduraClave = 12) OrElse (pbutton.Text = "13" AndAlso ArmaduraClave = 13) OrElse (pbutton.Text = "14" AndAlso ArmaduraClave = 14) Then
                CorretoIncorreto = "Correto"
                ValorCorreto += 1
            Else
                CorretoIncorreto = "Incorreto"
                ValorIncorreto += 1
            End If
            ValorTotal = ValorCorreto + ValorIncorreto
            ValorFinal = CInt(FormatNumber(((ValorCorreto / ValorTotal) * 100), 2))

            GerarAramduraDeClave()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click
        Me.WindowState = CType(1, FormWindowState) '1 é para minimizar
    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click
        Try

            Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click
        Try

            CheckBox1.Checked = True
            ValorCorreto = 0
            ValorIncorreto = 0
            ValorTotal = 0
            CorretoIncorreto = ""
            ValorInicial = 0
            GerarAramduraDeClave()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Label4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label4.Click
        Try

            CheckBox1.Checked = False
            ValorCorreto = 0
            ValorIncorreto = 0
            ValorTotal = 0
            CorretoIncorreto = ""
            ValorInicial = 0
            GerarAramduraDeClave()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub
End Class