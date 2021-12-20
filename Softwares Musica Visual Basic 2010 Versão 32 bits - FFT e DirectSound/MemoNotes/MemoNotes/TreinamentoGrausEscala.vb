Public Class TreinamentoGrausEscala

    Dim Nota(2), Valor As Integer

    Private Sub TreinamentoGrausEscala_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        GeraNota()
    End Sub

    Private Sub GeraNota()

        Try

            Nota(0) = 50
            Nota(2) = 50
            PictureBox2.Image = Nothing
            PictureBox2.Visible = False

            Do While (Nota(0) > 12 OrElse Nota(0) < 1)

                ' Create a byte array to hold the random value.
                Dim randomNumber(0) As Byte

                ' Create a new instance of the RNGCryptoServiceProvider.
                Dim Gen As New Security.Cryptography.RNGCryptoServiceProvider()

                ' Fill the array with a random value.
                Gen.GetBytes(randomNumber)

                ' Convert the byte to an integer value to make the modulus operation easier.
                Nota(0) = Int(Convert.ToInt32(randomNumber(0)))

                If Nota(0) = Nota(1) Then Nota(0) = 50
            Loop


            If Nota(0) = 1 Then
                PictureBox1.Image = My.Resources.C
            ElseIf Nota(0) = 2 Then
                SustenidoBemol()
                PictureBox2.Visible = True
                If Nota(2) = 1 Then
                    PictureBox1.Image = My.Resources.D
                    PictureBox2.Image = My.Resources.Bemol_grande
                ElseIf Nota(2) = 2 Then
                    PictureBox1.Image = My.Resources.C
                    PictureBox2.Image = My.Resources.Sustenido_grande
                End If
            ElseIf Nota(0) = 3 Then
                PictureBox1.Image = My.Resources.D
            ElseIf Nota(0) = 4 Then
                SustenidoBemol()
                PictureBox2.Visible = True
                If Nota(2) = 1 Then
                    PictureBox1.Image = My.Resources.E
                    PictureBox2.Image = My.Resources.Bemol_grande
                ElseIf Nota(2) = 2 Then
                    PictureBox1.Image = My.Resources.D
                    PictureBox2.Image = My.Resources.Sustenido_grande
                End If
            ElseIf Nota(0) = 5 Then
                PictureBox1.Image = My.Resources.E
            ElseIf Nota(0) = 6 Then
                PictureBox1.Image = My.Resources.F
            ElseIf Nota(0) = 7 Then
                SustenidoBemol()
                PictureBox2.Visible = True
                If Nota(2) = 1 Then
                    PictureBox1.Image = My.Resources.G
                    PictureBox2.Image = My.Resources.Bemol_grande
                ElseIf Nota(2) = 2 Then
                    PictureBox1.Image = My.Resources.F
                    PictureBox2.Image = My.Resources.Sustenido_grande
                End If
            ElseIf Nota(0) = 8 Then
                PictureBox1.Image = My.Resources.G
            ElseIf Nota(0) = 9 Then
                SustenidoBemol()
                PictureBox2.Visible = True
                If Nota(2) = 1 Then
                    PictureBox1.Image = My.Resources.A
                    PictureBox2.Image = My.Resources.Bemol_grande
                ElseIf Nota(2) = 2 Then
                    PictureBox1.Image = My.Resources.G
                    PictureBox2.Image = My.Resources.Sustenido_grande
                End If
            ElseIf Nota(0) = 10 Then
                PictureBox1.Image = My.Resources.A
            ElseIf Nota(0) = 11 Then
                SustenidoBemol()
                PictureBox2.Visible = True
                If Nota(2) = 1 Then
                    PictureBox1.Image = My.Resources.B
                    PictureBox2.Image = My.Resources.Bemol_grande
                ElseIf Nota(2) = 2 Then
                    PictureBox1.Image = My.Resources.A
                    PictureBox2.Image = My.Resources.Sustenido_grande
                End If
            ElseIf Nota(0) = 12 Then
                SustenidoBemol()
                PictureBox2.Visible = True
                If Nota(2) = 1 Then
                    PictureBox1.Image = My.Resources.C
                    PictureBox2.Image = My.Resources.Bemol_grande
                ElseIf Nota(2) = 2 Then
                    PictureBox1.Image = My.Resources.B
                End If
            End If

            Nota(1) = Nota(0)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub SustenidoBemol()

        Try

            Do While (Nota(2) > 2 OrElse Nota(2) < 1)

                ' Create a byte array to hold the random value.
                Dim randomNumber(0) As Byte

                ' Create a new instance of the RNGCryptoServiceProvider.
                Dim Gen As New Security.Cryptography.RNGCryptoServiceProvider()

                ' Fill the array with a random value.
                Gen.GetBytes(randomNumber)

                ' Convert the byte to an integer value to make the modulus operation easier.
                Nota(2) = Int(Convert.ToInt32(randomNumber(0)))
            Loop

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub VerificarAcerto()
        Try

            If Valor = Nota(0) Then
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

    Private Sub TreinamentoGrausEscala_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown, Button9.KeyDown, Button8.KeyDown, Button7.KeyDown, Button6.KeyDown, Button5.KeyDown, Button4.KeyDown, Button3.KeyDown, Button2.KeyDown, Button12.KeyDown, Button11.KeyDown, Button10.KeyDown, Button1.KeyDown

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
End Class