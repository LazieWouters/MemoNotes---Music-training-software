Public Class MemorizacaoNotasViolao

    Dim Nota(1), Valor As Integer

    Private Sub Desenhar(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint

        Try

            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
            Dim Cor As SolidBrush = New SolidBrush(Color.FromArgb(200, Color.Red))

            'setor 1
            If Nota(0) = 1 Then e.Graphics.FillEllipse(Cor, 55, 22, 18, 18)
            If Nota(0) = 2 Then e.Graphics.FillEllipse(Cor, 103, 22, 18, 18)
            If Nota(0) = 3 Then e.Graphics.FillEllipse(Cor, 151, 22, 18, 18)
            If Nota(0) = 4 Then e.Graphics.FillEllipse(Cor, 199, 22, 18, 18)

            If Nota(0) = 5 Then e.Graphics.FillEllipse(Cor, 55, 47, 18, 18)
            If Nota(0) = 6 Then e.Graphics.FillEllipse(Cor, 103, 47, 18, 18)
            If Nota(0) = 7 Then e.Graphics.FillEllipse(Cor, 151, 47, 18, 18)
            If Nota(0) = 8 Then e.Graphics.FillEllipse(Cor, 199, 47, 18, 18)

            If Nota(0) = 9 Then e.Graphics.FillEllipse(Cor, 55, 72, 18, 18)
            If Nota(0) = 10 Then e.Graphics.FillEllipse(Cor, 103, 72, 18, 18)
            If Nota(0) = 11 Then e.Graphics.FillEllipse(Cor, 151, 72, 18, 18)
            If Nota(0) = 12 Then e.Graphics.FillEllipse(Cor, 199, 72, 18, 18)

            If Nota(0) = 13 Then e.Graphics.FillEllipse(Cor, 55, 96, 18, 18)
            If Nota(0) = 14 Then e.Graphics.FillEllipse(Cor, 103, 96, 18, 18)
            If Nota(0) = 15 Then e.Graphics.FillEllipse(Cor, 151, 96, 18, 18)
            If Nota(0) = 16 Then e.Graphics.FillEllipse(Cor, 199, 96, 18, 18)

            If Nota(0) = 17 Then e.Graphics.FillEllipse(Cor, 55, 122, 18, 18)
            If Nota(0) = 18 Then e.Graphics.FillEllipse(Cor, 103, 122, 18, 18)
            If Nota(0) = 19 Then e.Graphics.FillEllipse(Cor, 151, 122, 18, 18)
            If Nota(0) = 20 Then e.Graphics.FillEllipse(Cor, 199, 122, 18, 18)

            If Nota(0) = 21 Then e.Graphics.FillEllipse(Cor, 55, 147, 18, 18)
            If Nota(0) = 22 Then e.Graphics.FillEllipse(Cor, 103, 147, 18, 18)
            If Nota(0) = 23 Then e.Graphics.FillEllipse(Cor, 151, 147, 18, 18)
            If Nota(0) = 24 Then e.Graphics.FillEllipse(Cor, 199, 147, 18, 18)




            'setor 2
            If Nota(0) = 25 Then e.Graphics.FillEllipse(Cor, 287, 22, 18, 18)
            If Nota(0) = 26 Then e.Graphics.FillEllipse(Cor, 335, 22, 18, 18)
            If Nota(0) = 27 Then e.Graphics.FillEllipse(Cor, 383, 22, 18, 18)
            If Nota(0) = 28 Then e.Graphics.FillEllipse(Cor, 431, 22, 18, 18)

            If Nota(0) = 29 Then e.Graphics.FillEllipse(Cor, 287, 47, 18, 18)
            If Nota(0) = 30 Then e.Graphics.FillEllipse(Cor, 335, 47, 18, 18)
            If Nota(0) = 31 Then e.Graphics.FillEllipse(Cor, 383, 47, 18, 18)
            If Nota(0) = 32 Then e.Graphics.FillEllipse(Cor, 431, 47, 18, 18)

            If Nota(0) = 33 Then e.Graphics.FillEllipse(Cor, 287, 72, 18, 18)
            If Nota(0) = 34 Then e.Graphics.FillEllipse(Cor, 335, 72, 18, 18)
            If Nota(0) = 35 Then e.Graphics.FillEllipse(Cor, 383, 72, 18, 18)
            If Nota(0) = 36 Then e.Graphics.FillEllipse(Cor, 431, 72, 18, 18)

            If Nota(0) = 37 Then e.Graphics.FillEllipse(Cor, 287, 96, 18, 18)
            If Nota(0) = 38 Then e.Graphics.FillEllipse(Cor, 335, 96, 18, 18)
            If Nota(0) = 39 Then e.Graphics.FillEllipse(Cor, 383, 96, 18, 18)
            If Nota(0) = 40 Then e.Graphics.FillEllipse(Cor, 431, 96, 18, 18)

            If Nota(0) = 41 Then e.Graphics.FillEllipse(Cor, 287, 122, 18, 18)
            If Nota(0) = 42 Then e.Graphics.FillEllipse(Cor, 335, 122, 18, 18)
            If Nota(0) = 43 Then e.Graphics.FillEllipse(Cor, 383, 122, 18, 18)
            If Nota(0) = 44 Then e.Graphics.FillEllipse(Cor, 431, 122, 18, 18)

            If Nota(0) = 45 Then e.Graphics.FillEllipse(Cor, 287, 147, 18, 18)
            If Nota(0) = 46 Then e.Graphics.FillEllipse(Cor, 335, 147, 18, 18)
            If Nota(0) = 47 Then e.Graphics.FillEllipse(Cor, 383, 147, 18, 18)
            If Nota(0) = 48 Then e.Graphics.FillEllipse(Cor, 431, 147, 18, 18)




            'setor 3
            If Nota(0) = 49 Then e.Graphics.FillEllipse(Cor, 519, 22, 18, 18)
            If Nota(0) = 50 Then e.Graphics.FillEllipse(Cor, 567, 22, 18, 18)
            If Nota(0) = 51 Then e.Graphics.FillEllipse(Cor, 615, 22, 18, 18)
            If Nota(0) = 52 Then e.Graphics.FillEllipse(Cor, 663, 22, 18, 18)

            If Nota(0) = 53 Then e.Graphics.FillEllipse(Cor, 519, 47, 18, 18)
            If Nota(0) = 54 Then e.Graphics.FillEllipse(Cor, 567, 47, 18, 18)
            If Nota(0) = 55 Then e.Graphics.FillEllipse(Cor, 615, 47, 18, 18)
            If Nota(0) = 56 Then e.Graphics.FillEllipse(Cor, 663, 47, 18, 18)

            If Nota(0) = 57 Then e.Graphics.FillEllipse(Cor, 519, 72, 18, 18)
            If Nota(0) = 58 Then e.Graphics.FillEllipse(Cor, 567, 72, 18, 18)
            If Nota(0) = 59 Then e.Graphics.FillEllipse(Cor, 615, 72, 18, 18)
            If Nota(0) = 60 Then e.Graphics.FillEllipse(Cor, 663, 72, 18, 18)

            If Nota(0) = 61 Then e.Graphics.FillEllipse(Cor, 519, 96, 18, 18)
            If Nota(0) = 62 Then e.Graphics.FillEllipse(Cor, 567, 96, 18, 18)
            If Nota(0) = 63 Then e.Graphics.FillEllipse(Cor, 615, 96, 18, 18)
            If Nota(0) = 64 Then e.Graphics.FillEllipse(Cor, 663, 96, 18, 18)

            If Nota(0) = 65 Then e.Graphics.FillEllipse(Cor, 519, 122, 18, 18)
            If Nota(0) = 66 Then e.Graphics.FillEllipse(Cor, 567, 122, 18, 18)
            If Nota(0) = 67 Then e.Graphics.FillEllipse(Cor, 615, 122, 18, 18)
            If Nota(0) = 68 Then e.Graphics.FillEllipse(Cor, 663, 122, 18, 18)

            If Nota(0) = 69 Then e.Graphics.FillEllipse(Cor, 519, 147, 18, 18)
            If Nota(0) = 70 Then e.Graphics.FillEllipse(Cor, 567, 147, 18, 18)
            If Nota(0) = 71 Then e.Graphics.FillEllipse(Cor, 615, 147, 18, 18)
            If Nota(0) = 72 Then e.Graphics.FillEllipse(Cor, 663, 147, 18, 18)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub GeraNota()

        Try

            Nota(0) = 100


            Do While (Nota(0) > 72 OrElse Nota(0) < 1)

                ' Create a byte array to hold the random value.
                Dim randomNumber(0) As Byte

                ' Create a new instance of the RNGCryptoServiceProvider.
                Dim Gen As New Security.Cryptography.RNGCryptoServiceProvider()

                ' Fill the array with a random value.
                Gen.GetBytes(randomNumber)

                ' Convert the byte to an integer value to make the modulus operation easier.
                Nota(0) = Int(Convert.ToInt32(randomNumber(0)))

                If Nota(0) = Nota(1) Then Nota(0) = 100

                If (Nota(0) >= 1 AndAlso Nota(0) <= 24) AndAlso Not CheckBox1.Checked Then
                    Nota(0) = 100
                ElseIf (Nota(0) >= 25 AndAlso Nota(0) <= 48) AndAlso Not CheckBox2.Checked Then
                    Nota(0) = 100
                ElseIf (Nota(0) >= 49 AndAlso Nota(0) <= 72) AndAlso Not CheckBox3.Checked Then
                    Nota(0) = 100
                End If

                If CheckBox4.Checked AndAlso (Nota(0) = 2 OrElse Nota(0) = 4 OrElse Nota(0) = 6 OrElse Nota(0) = 8 OrElse Nota(0) = 9 OrElse Nota(0) = 11 OrElse Nota(0) = 13 OrElse Nota(0) = 16 _
                                           OrElse Nota(0) = 17 OrElse Nota(0) = 20 OrElse Nota(0) = 22 OrElse Nota(0) = 24 OrElse Nota(0) = 26 OrElse Nota(0) = 31 OrElse Nota(0) = 34 OrElse Nota(0) = 36 _
                                            OrElse Nota(0) = 38 OrElse Nota(0) = 40 OrElse Nota(0) = 42 OrElse Nota(0) = 46 OrElse Nota(0) = 49 OrElse Nota(0) = 51 OrElse Nota(0) = 53 OrElse Nota(0) = 55 _
                                             OrElse Nota(0) = 59 OrElse Nota(0) = 63 OrElse Nota(0) = 65 OrElse Nota(0) = 67 OrElse Nota(0) = 69 OrElse Nota(0) = 71) Then
                    Nota(0) = 100
                End If

            Loop



            Nota(1) = Nota(0)

            Dim Rect1 As New Rectangle(52, 16, 172, 162)
            Me.Invalidate(Rect1)

            Dim Rect2 As New Rectangle(283, 16, 172, 162)
            Me.Invalidate(Rect2)

            Dim Rect3 As New Rectangle(515, 16, 172, 162)
            Me.Invalidate(Rect3)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub TreinamentoGrausEscala_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Button9.KeyDown, Button8.KeyDown, Button7.KeyDown, Button6.KeyDown, Button5.KeyDown, Button4.KeyDown, Button3.KeyDown, Button2.KeyDown, Button12.KeyDown, Button11.KeyDown, Button10.KeyDown, Button1.KeyDown, CheckBox1.KeyDown, CheckBox2.KeyDown, CheckBox3.KeyDown, CheckBox4.KeyDown

        Try

            If e.KeyCode = Keys.F1 OrElse e.KeyCode = Keys.NumPad1 Then
                Valor = 1
            ElseIf e.KeyCode = Keys.F2 OrElse e.KeyCode = Keys.NumPad2 Then
                Valor = 2
            ElseIf e.KeyCode = Keys.F3 OrElse e.KeyCode = Keys.NumPad3 Then
                Valor = 3
            ElseIf e.KeyCode = Keys.F4 OrElse e.KeyCode = Keys.NumPad4 Then
                Valor = 4
            ElseIf e.KeyCode = Keys.F5 OrElse e.KeyCode = Keys.NumPad5 Then
                Valor = 5
            ElseIf e.KeyCode = Keys.F6 OrElse e.KeyCode = Keys.NumPad6 Then
                Valor = 6
            ElseIf e.KeyCode = Keys.F7 OrElse e.KeyCode = Keys.NumPad7 Then
                Valor = 7
            ElseIf e.KeyCode = Keys.F8 OrElse e.KeyCode = Keys.NumPad8 Then
                Valor = 8
            ElseIf e.KeyCode = Keys.F9 OrElse e.KeyCode = Keys.NumPad9 Then
                Valor = 9
            ElseIf e.KeyCode = Keys.F10 OrElse e.KeyCode = Keys.Divide Then
                Valor = 10
            ElseIf e.KeyCode = Keys.F11 OrElse e.KeyCode = Keys.Multiply Then
                Valor = 11
            ElseIf e.KeyCode = Keys.F12 OrElse e.KeyCode = Keys.Subtract Then
                Valor = 12
            End If

            VerificarAcerto()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub VerificarAcerto()

        Try

            If (Valor = 1 AndAlso (Nota(0) = 5 OrElse Nota(0) = 19 OrElse Nota(0) = 28 OrElse Nota(0) = 33 OrElse Nota(0) = 48 OrElse Nota(0) = 62)) OrElse _
                (Valor = 2 AndAlso (Nota(0) = 6 OrElse Nota(0) = 20 OrElse Nota(0) = 34 OrElse Nota(0) = 49 OrElse Nota(0) = 63 OrElse Nota(0) = 69) OrElse _
                (Valor = 3 AndAlso (Nota(0) = 7 OrElse Nota(0) = 35 OrElse Nota(0) = 41 OrElse Nota(0) = 50 OrElse Nota(0) = 64 OrElse Nota(0) = 70)) OrElse _
                (Valor = 4 AndAlso (Nota(0) = 8 OrElse Nota(0) = 13 OrElse Nota(0) = 36 OrElse Nota(0) = 42 OrElse Nota(0) = 51 OrElse Nota(0) = 71)) OrElse _
                (Valor = 5 AndAlso (Nota(0) = 14 OrElse Nota(0) = 29 OrElse Nota(0) = 43 OrElse Nota(0) = 52 OrElse Nota(0) = 57 OrElse Nota(0) = 72)) OrElse _
                (Valor = 6 AndAlso (Nota(0) = 1 OrElse Nota(0) = 15 OrElse Nota(0) = 21 OrElse Nota(0) = 30 OrElse Nota(0) = 44 OrElse Nota(0) = 58)) OrElse _
                (Valor = 7 AndAlso (Nota(0) = 2 OrElse Nota(0) = 16 OrElse Nota(0) = 22 OrElse Nota(0) = 31 OrElse Nota(0) = 59 OrElse Nota(0) = 65)) OrElse _
                (Valor = 8 AndAlso (Nota(0) = 3 OrElse Nota(0) = 23 OrElse Nota(0) = 32 OrElse Nota(0) = 37 OrElse Nota(0) = 60 OrElse Nota(0) = 66)) OrElse _
                (Valor = 9 AndAlso (Nota(0) = 4 OrElse Nota(0) = 9 OrElse Nota(0) = 24 OrElse Nota(0) = 38 OrElse Nota(0) = 53 OrElse Nota(0) = 67)) OrElse _
                (Valor = 10 AndAlso (Nota(0) = 10 OrElse Nota(0) = 25 OrElse Nota(0) = 39 OrElse Nota(0) = 45 OrElse Nota(0) = 54 OrElse Nota(0) = 68)) OrElse _
                (Valor = 11 AndAlso (Nota(0) = 11 OrElse Nota(0) = 17 OrElse Nota(0) = 26 OrElse Nota(0) = 40 OrElse Nota(0) = 46 OrElse Nota(0) = 55)) OrElse _
                (Valor = 12 AndAlso (Nota(0) = 12 OrElse Nota(0) = 18 OrElse Nota(0) = 27 OrElse Nota(0) = 47 OrElse Nota(0) = 56 OrElse Nota(0) = 61))) Then
                Label2.Text = "Correto"
                Label2.ForeColor = Color.DarkGreen

            Else
                Label2.Text = "Incorreto"
                Label2.ForeColor = Color.Red
            End If
            GeraNota()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Valor = 1
        VerificarAcerto()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Valor = 2
        VerificarAcerto()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Valor = 3
        VerificarAcerto()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Valor = 4
        VerificarAcerto()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Valor = 5
        VerificarAcerto()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Valor = 6
        VerificarAcerto()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Valor = 7
        VerificarAcerto()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Valor = 8
        VerificarAcerto()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Valor = 9
        VerificarAcerto()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Valor = 10
        VerificarAcerto()
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Valor = 11
        VerificarAcerto()
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Valor = 12
        VerificarAcerto()
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged, CheckBox2.CheckedChanged, CheckBox1.CheckedChanged
        If Not CheckBox1.Checked And Not CheckBox2.Checked And Not CheckBox3.Checked Then CheckBox1.Checked = True
    End Sub

    Private Sub MemorizacaoNotasViolao_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        SalvaSettings()
    End Sub

    Private Sub SalvaSettings()
        Try

            My.Settings.NovoValorViolaoSetores(0) = CheckBox1.Checked
            My.Settings.NovoValorViolaoSetores(1) = CheckBox2.Checked
            My.Settings.NovoValorViolaoSetores(2) = CheckBox3.Checked
            My.Settings.NovoValorViolaoSomenteNotasNaturais = CheckBox4.Checked

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub MemorizacaoNotasViolao_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

            CheckBox1.Checked = My.Settings.NovoValorViolaoSetores(0)
            CheckBox2.Checked = My.Settings.NovoValorViolaoSetores(1)
            CheckBox3.Checked = My.Settings.NovoValorViolaoSetores(2)
            CheckBox4.Checked = My.Settings.NovoValorViolaoSomenteNotasNaturais

            GeraNota()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub
End Class