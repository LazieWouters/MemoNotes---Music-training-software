Option Strict Off
Option Explicit On
Imports System.Math

Public Class TreinamentoIntervalos

    Dim ContadorIntervalo, QuantidadeIntervalos, Nota(20), ValorTickTimer, ValorTickTimer2, spin, NotaMidi As Integer
    Dim Sequencia, Direção, Intervalo, NomeIntervalo(20) As String
    Dim Cor1 As SolidBrush = New SolidBrush(Color.FromArgb(50, 0, 0, 255))
    'Dim Img As Bitmap = My.Resources.ReconhecimentoIntervalos


    Private Sub GeraNota()

        Try

            Array.Clear(Nota, 0, 20)
            Array.Clear(NomeIntervalo, 0, 20)

            QuantidadeIntervalos = CInt(QuantidadeDeIntervalos.Text) + 1

            ContadorIntervalo = 1
            Do While ContadorIntervalo <= QuantidadeIntervalos
                If ContadorIntervalo = 1 Then
                    Do While Nota(ContadorIntervalo) <= 30 OrElse Nota(ContadorIntervalo) >= 60
                        NúmeroAleatório()
                    Loop
                Else
                    Do While Nota(ContadorIntervalo) <= Nota(ContadorIntervalo - 1) - 12 OrElse Nota(ContadorIntervalo) >= Nota(ContadorIntervalo - 1) + 12
                        NúmeroAleatório()

                        If (Not Ascendente.Checked AndAlso Nota(ContadorIntervalo) > Nota(ContadorIntervalo - 1)) OrElse
                        (Not Descendente.Checked AndAlso Nota(ContadorIntervalo) < Nota(ContadorIntervalo - 1)) Then Nota(ContadorIntervalo) = 0

                        If (Not SegundaMenor.Checked AndAlso Abs(Nota(ContadorIntervalo) - Nota(ContadorIntervalo - 1)) = 1) OrElse
                        (Not SegundaMaior.Checked AndAlso Abs(Nota(ContadorIntervalo) - Nota(ContadorIntervalo - 1)) = 2) OrElse
                        (Not TerçaMenor.Checked AndAlso Abs(Nota(ContadorIntervalo) - Nota(ContadorIntervalo - 1)) = 3) OrElse
                        (Not TerçaMaior.Checked AndAlso Abs(Nota(ContadorIntervalo) - Nota(ContadorIntervalo - 1)) = 4) OrElse
                        (Not QuartaJusta.Checked AndAlso Abs(Nota(ContadorIntervalo) - Nota(ContadorIntervalo - 1)) = 5) OrElse
                        (Not QuintaDiminuida.Checked AndAlso Abs(Nota(ContadorIntervalo) - Nota(ContadorIntervalo - 1)) = 6) OrElse
                        (Not QuintaJusta.Checked AndAlso Abs(Nota(ContadorIntervalo) - Nota(ContadorIntervalo - 1)) = 7) OrElse
                        (Not SextaMenor.Checked AndAlso Abs(Nota(ContadorIntervalo) - Nota(ContadorIntervalo - 1)) = 8) OrElse
                        (Not SextaMaior.Checked AndAlso Abs(Nota(ContadorIntervalo) - Nota(ContadorIntervalo - 1)) = 9) OrElse
                        (Not SétimaMenor.Checked AndAlso Abs(Nota(ContadorIntervalo) - Nota(ContadorIntervalo - 1)) = 10) OrElse
                        (Not SétimaMaior.Checked AndAlso Abs(Nota(ContadorIntervalo) - Nota(ContadorIntervalo - 1)) = 11) OrElse
                        (Not Oitava.Checked AndAlso Abs(Nota(ContadorIntervalo) - Nota(ContadorIntervalo - 1)) = 12) Then Nota(ContadorIntervalo) = 0

                        If Nota(ContadorIntervalo) <= 10 OrElse Nota(ContadorIntervalo) >= 80 Then Nota(ContadorIntervalo) = 0

                        If Nota(ContadorIntervalo) <> 0 Then
                            If Nota(ContadorIntervalo) > Nota(ContadorIntervalo - 1) Then
                                Direção = "A]"
                            ElseIf Nota(ContadorIntervalo) < Nota(ContadorIntervalo - 1) Then
                                Direção = "D]"
                            Else
                                Direção = "]"
                            End If


                            If Abs(Nota(ContadorIntervalo) - Nota(ContadorIntervalo - 1)) = 0 Then
                                Intervalo = "[U"
                            ElseIf Abs(Nota(ContadorIntervalo) - Nota(ContadorIntervalo - 1)) = 1 Then
                                Intervalo = "[2m "
                            ElseIf Abs(Nota(ContadorIntervalo) - Nota(ContadorIntervalo - 1)) = 2 Then
                                Intervalo = "[2M "
                            ElseIf Abs(Nota(ContadorIntervalo) - Nota(ContadorIntervalo - 1)) = 3 Then
                                Intervalo = "[3m "
                            ElseIf Abs(Nota(ContadorIntervalo) - Nota(ContadorIntervalo - 1)) = 4 Then
                                Intervalo = "[3M "
                            ElseIf Abs(Nota(ContadorIntervalo) - Nota(ContadorIntervalo - 1)) = 5 Then
                                Intervalo = "[4J "
                            ElseIf Abs(Nota(ContadorIntervalo) - Nota(ContadorIntervalo - 1)) = 6 Then
                                Intervalo = "[4a "
                            ElseIf Abs(Nota(ContadorIntervalo) - Nota(ContadorIntervalo - 1)) = 7 Then
                                Intervalo = "[5J "
                            ElseIf Abs(Nota(ContadorIntervalo) - Nota(ContadorIntervalo - 1)) = 8 Then
                                Intervalo = "[6m "
                            ElseIf Abs(Nota(ContadorIntervalo) - Nota(ContadorIntervalo - 1)) = 9 Then
                                Intervalo = "[6M "
                            ElseIf Abs(Nota(ContadorIntervalo) - Nota(ContadorIntervalo - 1)) = 10 Then
                                Intervalo = "[7m "
                            ElseIf Abs(Nota(ContadorIntervalo) - Nota(ContadorIntervalo - 1)) = 11 Then
                                Intervalo = "[7M "
                            ElseIf Abs(Nota(ContadorIntervalo) - Nota(ContadorIntervalo - 1)) = 12 Then
                                Intervalo = "[8P "
                            End If

                            NomeIntervalo(ContadorIntervalo) = Intervalo & Direção
                        End If

                    Loop
                End If

                ContadorIntervalo += 1

            Loop

            AtivaTimer()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub NúmeroAleatório()
        Try

            ' Create a byte array to hold the random value.
            Dim randomNumber(0) As Byte

            ' Create a new instance of the RNGCryptoServiceProvider.
            Dim Gen As New Security.Cryptography.RNGCryptoServiceProvider()

            ' Fill the array with a random value.
            Gen.GetBytes(randomNumber)

            ' Convert the byte to an integer value to make the modulus operation easier.
            Nota(ContadorIntervalo) = Int(Convert.ToInt32(randomNumber(0)))

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Try

            If ValorTickTimer <= QuantidadeIntervalos Then
                Try
                    If NotaMidi <> 1000 Then MidiPlayer.Play(New NoteOff(0, 1, CByte(NotaMidi), CByte(NumericUpDown2.Value)))
                Catch ex As Exception

                    'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
                Finally
                    'esta parte do bloco é executada independentemente se algum erro acontecer ou não

                End Try

                Sequencia = Sequencia & NomeIntervalo(ValorTickTimer) & "   "
                'Label4.Text = Label4.Text & "?   "
                Label4.Text = ""
                Label4.Text = Label4.Text & Sequencia & "  "
                'NomeAcorde(100) = NomeAcorde(ValorTickTimer)
                If ComboBox1.Text = "Piano Bösendorfer" Then
                    mp3 = "Sons\Notas\Piano Bösendorfer\P" & Nota(ValorTickTimer) & ".mp3"
                    'TocaSom()
                ElseIf ComboBox1.Text = "Strings (Trio)" Then
                    mp3 = "Sons\Notas\Strings (Trio)\P" & Nota(ValorTickTimer) & ".mp3"
                    'TocaSom()
                Else
                    Timer3.Enabled = False
                    Try
                        MidiPlayer.Play(New ProgramChange(0, 1, CType(ComboBox1.Text.Substring(0, 3), GeneralMidiInstruments)))
                        MidiPlayer.Play(New NoteOn(0, 1, CByte(Nota(ValorTickTimer) + 20), CByte(NumericUpDown2.Value)))
                    Catch ex As Exception

                        'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
                    Finally
                        'esta parte do bloco é executada independentemente se algum erro acontecer ou não

                    End Try

                    NotaMidi = Nota(ValorTickTimer) + 20
                    Timer3.Enabled = True
                End If
            Else
                Timer1.Enabled = False
            End If
            ValorTickTimer += 1

            If Timer1.Interval = 1 Then Timer1.Interval = CInt(((CDbl(Minutos.Text) * 60) + CDbl(Segundos.Text) + (CDbl(Decimos.Text) / 100)) * 1000) : Timer3.Interval = Timer1.Interval

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click
        If CInt(QuantidadeDeIntervalos.Text) < 10 Then QuantidadeDeIntervalos.Text = CStr(CDbl(QuantidadeDeIntervalos.Text) + 1)
        AtivarBotoesIntervalos()
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        If CInt(QuantidadeDeIntervalos.Text) > 1 Then QuantidadeDeIntervalos.Text = CStr(CDbl(QuantidadeDeIntervalos.Text) - 1)
        AtivarBotoesIntervalos()
    End Sub

    Private Sub AtivarBotoesIntervalos() Handles Me.Load
        Try

            If QuantidadeDeIntervalos.Text > 0 Then Button27.Enabled = True Else Button27.Enabled = False
            If QuantidadeDeIntervalos.Text > 1 Then Button28.Enabled = True Else Button28.Enabled = False
            If QuantidadeDeIntervalos.Text > 2 Then Button29.Enabled = True Else Button29.Enabled = False
            If QuantidadeDeIntervalos.Text > 3 Then Button30.Enabled = True Else Button30.Enabled = False
            If QuantidadeDeIntervalos.Text > 4 Then Button31.Enabled = True Else Button31.Enabled = False
            If QuantidadeDeIntervalos.Text > 5 Then Button32.Enabled = True Else Button32.Enabled = False
            If QuantidadeDeIntervalos.Text > 6 Then Button33.Enabled = True Else Button33.Enabled = False
            If QuantidadeDeIntervalos.Text > 7 Then Button34.Enabled = True Else Button34.Enabled = False
            If QuantidadeDeIntervalos.Text > 8 Then Button35.Enabled = True Else Button35.Enabled = False
            If QuantidadeDeIntervalos.Text > 9 Then Button36.Enabled = True Else Button36.Enabled = False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub PictureBox26_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox26.MouseDown
        If e.Button = MouseButtons.Left Then
            spinUP()
            spin = 1
            Timer5.Enabled = True
        End If
    End Sub

    Private Sub PictureBox26_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox26.MouseUp
        Timer5.Enabled = False
    End Sub

    Private Sub PictureBox27_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox27.MouseDown
        If e.Button = MouseButtons.Left Then
            spinDOWN()
            spin = 2
            Timer5.Enabled = True
        End If
    End Sub

    Private Sub PictureBox27_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox27.MouseUp
        Timer5.Enabled = False
    End Sub

    Private Sub Timer5_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer5.Tick
        If spin = 1 Then
            spinUP()
        Else
            spinDOWN()
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
            Timer3.Interval = Timer1.Interval

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
            Timer3.Interval = Timer1.Interval

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub Minutos_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Minutos.Leave

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
            Timer3.Interval = Timer1.Interval

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub Segundos_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Segundos.Leave

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
            Timer3.Interval = Timer1.Interval

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub Decimos_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Decimos.Leave

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
            Timer3.Interval = Timer1.Interval

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub Button25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button25.Click
        GeraNota()
    End Sub

    Private Sub TreinamentoIntervalos_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            MidiPlayer.CloseMidi()
        Catch ex As Exception
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

            SalvaSettings()
            'CloseClips()
        End Try
    End Sub

    Private Sub SalvaSettings()
        Try

            My.Settings.NovoValorIntervalos(0) = SegundaMenor.Checked
            My.Settings.NovoValorIntervalos(1) = SegundaMaior.Checked
            My.Settings.NovoValorIntervalos(2) = TerçaMenor.Checked
            My.Settings.NovoValorIntervalos(3) = TerçaMaior.Checked
            My.Settings.NovoValorIntervalos(4) = QuartaJusta.Checked
            My.Settings.NovoValorIntervalos(5) = QuintaDiminuida.Checked
            My.Settings.NovoValorIntervalos(6) = QuintaJusta.Checked
            My.Settings.NovoValorIntervalos(7) = SextaMenor.Checked
            My.Settings.NovoValorIntervalos(8) = SextaMaior.Checked
            My.Settings.NovoValorIntervalos(9) = SétimaMenor.Checked
            My.Settings.NovoValorIntervalos(10) = SétimaMaior.Checked
            My.Settings.NovoValorIntervalos(11) = Oitava.Checked

            My.Settings.NovoValorAscendente = Ascendente.Checked
            My.Settings.NovoValorDescendente = Descendente.Checked

            My.Settings.NovoValorTimers(9) = Minutos.Text
            My.Settings.NovoValorTimers(10) = Segundos.Text
            My.Settings.NovoValorTimers(11) = Decimos.Text

            My.Settings.NovoValorQtdeIntervalos = QuantidadeDeIntervalos.Text

            My.Settings.NovoValorInstrumentoMusical(2) = ComboBox1.Text

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub TreinamentoIntervalos_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            NotaMidi = 1000

            MidiPlayer.OpenMidi()
        Catch ex As Exception
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

            'Img.MakeTransparent(Img.GetPixel(10, 10))
            'Me.BackgroundImage = Img
            'Me.TransparencyKey = Img.GetPixel(10, 10)
            Me.BringToFront()

            SegundaMenor.Checked = My.Settings.NovoValorIntervalos(0)
            SegundaMaior.Checked = My.Settings.NovoValorIntervalos(1)
            TerçaMenor.Checked = My.Settings.NovoValorIntervalos(2)
            TerçaMaior.Checked = My.Settings.NovoValorIntervalos(3)
            QuartaJusta.Checked = My.Settings.NovoValorIntervalos(4)
            QuintaDiminuida.Checked = My.Settings.NovoValorIntervalos(5)
            QuintaJusta.Checked = My.Settings.NovoValorIntervalos(6)
            SextaMenor.Checked = My.Settings.NovoValorIntervalos(7)
            SextaMaior.Checked = My.Settings.NovoValorIntervalos(8)
            SétimaMenor.Checked = My.Settings.NovoValorIntervalos(9)
            SétimaMaior.Checked = My.Settings.NovoValorIntervalos(10)
            Oitava.Checked = My.Settings.NovoValorIntervalos(11)

            Ascendente.Checked = My.Settings.NovoValorAscendente
            Descendente.Checked = My.Settings.NovoValorDescendente

            Minutos.Text = My.Settings.NovoValorTimers(9)
            Segundos.Text = My.Settings.NovoValorTimers(10)
            Decimos.Text = My.Settings.NovoValorTimers(11)

            QuantidadeDeIntervalos.Text = My.Settings.NovoValorQtdeIntervalos

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

            ComboBox1.Text = My.Settings.NovoValorInstrumentoMusical(2)

            If ComboBox1.Text = "" Then ComboBox1.Text = "000 - Acoustic Grand Piano"

            Tocanota = 1
        End Try

    End Sub

    Private Sub Button26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button26.Click
        AtivaTimer()
    End Sub

    Private Sub AtivaTimer()
        Try

            ValorTickTimer = 1
            Label4.Text = ""
            Sequencia = ""
            Timer1.Interval = 1
            Timer1.Enabled = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub ab_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown, Label4.MouseDown, Label3.MouseDown
        VGa = Me.MousePosition.X - Me.Location.X
        VGb = Me.MousePosition.Y - Me.Location.Y
    End Sub

    Private Sub ab_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove, Label4.MouseMove, Label3.MouseMove
        If e.Button = MouseButtons.Left Then
            VGnewPoint = Me.MousePosition
            VGnewPoint.X -= VGa
            VGnewPoint.Y -= VGb
            Me.Location = VGnewPoint
        End If
    End Sub

    Private Sub SegundaMenor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2menor.Click
        SegundaMenor.Checked = Not SegundaMenor.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub SegundaMaior_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2maior.Click
        SegundaMaior.Checked = Not SegundaMaior.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub TerçaMenor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox3menor.Click
        TerçaMenor.Checked = Not TerçaMenor.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub TerçaMaior_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox3maior.Click
        TerçaMaior.Checked = Not TerçaMaior.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub QuartaJusta_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox4justa.Click
        QuartaJusta.Checked = Not QuartaJusta.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub QuintaDiminuida_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox5diminuta.Click
        QuintaDiminuida.Checked = Not QuintaDiminuida.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub QuintaJusta_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox5justa.Click
        QuintaJusta.Checked = Not QuintaJusta.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub SextaMenor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox6menor.Click
        SextaMenor.Checked = Not SextaMenor.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub SextaMaior_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox6maior.Click
        SextaMaior.Checked = Not SextaMaior.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub SétimaMenor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox7menor.Click
        SétimaMenor.Checked = Not SétimaMenor.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub SétimaMaior_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox7maior.Click
        SétimaMaior.Checked = Not SétimaMaior.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub Oitava_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox8justa.Click
        Oitava.Checked = Not Oitava.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub SegundaMenor_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox2menor.Paint
        If SegundaMenor.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub SegundaMaior_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox2maior.Paint
        If SegundaMaior.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub TerçaMenor_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox3menor.Paint
        If TerçaMenor.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub TerçaMaior_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox3maior.Paint
        If TerçaMaior.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub QuartaJusta_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox4justa.Paint
        If QuartaJusta.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub QuintaDiminuida_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox5diminuta.Paint
        If QuintaDiminuida.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub QuintaJusta_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox5justa.Paint
        If QuintaJusta.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub SextaMenor_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox6menor.Paint
        If SextaMenor.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub SextaMaior_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox6maior.Paint
        If SextaMaior.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub SétimaMenor_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox7menor.Paint
        If SétimaMenor.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub SétimaMaior_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox7maior.Paint
        If SétimaMaior.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub Oitava_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox8justa.Paint
        If Oitava.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub ChecarFigurasNecessárias()



    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick

        Try

            If ValorTickTimer <= ValorTickTimer2 + 1 Then
                Try
                    If NotaMidi <> 1000 Then MidiPlayer.Play(New NoteOff(0, 1, CByte(NotaMidi), CByte(NumericUpDown2.Value)))
                Catch ex As Exception

                    'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
                Finally
                    'esta parte do bloco é executada independentemente se algum erro acontecer ou não

                End Try

                If ComboBox1.Text = "Piano Bösendorfer" Then
                    mp3 = "Sons\Notas\Piano Bösendorfer\P" & Nota(ValorTickTimer) & ".mp3"
                    'TocaSom()
                ElseIf ComboBox1.Text = "Strings (Trio)" Then
                    mp3 = "Sons\Notas\Strings (Trio)\P" & Nota(ValorTickTimer) & ".mp3"
                    'TocaSom()
                Else
                    Timer3.Enabled = False
                    Try
                        MidiPlayer.Play(New ProgramChange(0, 1, CType(ComboBox1.Text.Substring(0, 3), GeneralMidiInstruments)))
                        MidiPlayer.Play(New NoteOn(0, 1, CByte(Nota(ValorTickTimer) + 20), CByte(NumericUpDown2.Value)))
                    Catch ex As Exception

                        'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
                    Finally
                        'esta parte do bloco é executada independentemente se algum erro acontecer ou não

                    End Try

                    NotaMidi = Nota(ValorTickTimer) + 20
                    Timer3.Enabled = True
                End If
            Else
                Timer2.Enabled = False
            End If
            ValorTickTimer += 1

            If Timer2.Interval = 1 Then Timer2.Interval = CInt(((CDbl(Minutos.Text) * 60) + CDbl(Segundos.Text) + (CDbl(Decimos.Text) / 100)) * 1000) : Timer3.Interval = Timer2.Interval

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

    Private Sub Button27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button27.Click
        Try

            ValorTickTimer = 1
            ValorTickTimer2 = 1
            Timer2.Interval = 1
            Timer2.Enabled = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Button28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button28.Click
        Try

            ValorTickTimer = 2
            ValorTickTimer2 = 2
            Timer2.Interval = 1
            Timer2.Enabled = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Button29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button29.Click
        Try

            ValorTickTimer = 3
            ValorTickTimer2 = 3
            Timer2.Interval = 1
            Timer2.Enabled = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Button30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button30.Click
        Try

            ValorTickTimer = 4
            ValorTickTimer2 = 4
            Timer2.Interval = 1
            Timer2.Enabled = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Button31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button31.Click
        Try

            ValorTickTimer = 5
            ValorTickTimer2 = 5
            Timer2.Interval = 1
            Timer2.Enabled = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Button32_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button32.Click
        Try

            ValorTickTimer = 6
            ValorTickTimer2 = 6
            Timer2.Interval = 1
            Timer2.Enabled = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Button33_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button33.Click
        Try

            ValorTickTimer = 7
            ValorTickTimer2 = 7
            Timer2.Interval = 1
            Timer2.Enabled = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Button34_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button34.Click
        Try

            ValorTickTimer = 8
            ValorTickTimer2 = 8
            Timer2.Interval = 1
            Timer2.Enabled = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Button35_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button35.Click
        Try

            ValorTickTimer = 9
            ValorTickTimer2 = 9
            Timer2.Interval = 1
            Timer2.Enabled = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Button36_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button36.Click
        Try

            ValorTickTimer = 10
            ValorTickTimer2 = 10
            Timer2.Interval = 1
            Timer2.Enabled = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Button37_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button37.Click
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

End Class