
Public Class ExercicioParaMemorizarEscalas

    Dim a, b, c, d, f, g, h, j, k, l, ValorLeft As Integer
    Dim Nota, Nota2 As String
    Dim Deletar As Boolean
    Dim Borda As Pen = New Pen(Color.FromArgb(0, 0, 0), 4)
    Dim Borda2 As Pen = New Pen(Color.FromArgb(0, 0, 0), 1)
    Dim Fonte As New Font("Arial", 9, FontStyle.Regular, GraphicsUnit.Pixel)
    Dim CorFonte As SolidBrush = New SolidBrush(Color.FromArgb(50, 50, 50))

    Dim EscalaMaior(,) As String = {{"C", "D", "E", "F", "G", "A", "B"}, _
                                       {"G", "A", "B", "C", "D", "E", "F#"}, _
                                       {"D", "E", "F#", "G", "A", "B", "C#"}, _
                                       {"A", "B", "C#", "D", "E", "F#", "G#"}, _
                                       {"E", "F#", "G#", "A", "B", "C#", "D#"}, _
                                       {"B", "C#", "D#", "E", "F#", "G#", "A#"}, _
                                       {"Cb", "Db", "Eb", "Fb", "Gb", "Ab", "Bb"}, _
                                       {"F#", "G#", "A#", "B", "C#", "D#", "E#"}, _
                                       {"Gb", "Ab", "Bb", "Cb", "Db", "Eb", "F"}, _
                                       {"C#", "D#", "E#", "F#", "G#", "A#", "B#"}, _
                                       {"Db", "Eb", "F", "Gb", "Ab", "Bb", "C"}, _
                                       {"Ab", "Bb", "C", "Db", "Eb", "F", "G"}, _
                                       {"Eb", "F", "G", "Ab", "Bb", "C", "D"}, _
                                       {"Bb", "C", "D", "Eb", "F", "G", "A"}, _
                                       {"F", "G", "A", "Bb", "C", "D", "E"}, _
                                       {"G#", "A#", "B#", "C#", "D#", "E#", "F##"}, _
                                       {"Bbb", "Cb", "Db", "Ebb", "Fb", "Gb", "Ab"}, _
                                       {"Abb", "Bbb", "Cb", "D", "Ebb", "Fb", "Gb"}, _
                                       {"Fb", "Gb", "Ab", "Bbb", "Cb", "Db", "Eb"}, _
                                       {"Ebb", "Fb", "Gb", "Abb", "Bbb", "Cb", "Db"}, _
                                       {"D#", "E#", "F##", "G#", "A#", "B#", "C##"}, _
                                       {"E#", "F##", "G##", "A#", "B#", "C##", "D##"}, _
                                       {"A#", "B#", "C##", "D#", "E#", "F##", "G##"}, _
                                       {"B#", "C##", "D##", "E#", "F##", "G##", "A##"}, _
                                       {"F##", "G##", "A##", "B#", "C##", "D##", "E##"}, _
                                       {"Gbb", "Abb", "Bbb", "Cbb", "Dbb", "Ebb", "Fb"}, _
                                       {"Dbb", "Ebb", "Fb", "Gbb", "Abb", "Bbb", "Cb"}, _
                                       {"G##", "A##", "B##", "C##", "D##", "E##", "F###"}, _
                                       {"C##", "D##", "E##", "F##", "G##", "A##", "B##"}, _
                                       {"D##", "E##", "F###", "G##", "A##", "B##", "C###"}}

    Dim EscalaMenorNatural(,) As String = {{"C", "D", "Eb", "F", "G", "Ab", "Bb"}, _
                                       {"G", "A", "Bb", "C", "D", "Eb", "F"}, _
                                       {"D", "E", "F", "G", "A", "Bb", "C"}, _
                                       {"A", "B", "C", "D", "E", "F", "G"}, _
                                       {"E", "F#", "G", "A", "B", "C", "D"}, _
                                       {"B", "C#", "D", "E", "F#", "G", "A"}, _
                                       {"Cb", "Db", "Ebb", "Fb", "Gb", "Abb", "Bbb"}, _
                                       {"F#", "G#", "A", "B", "C#", "D", "E"}, _
                                       {"Gb", "Ab", "Bbb", "Cb", "Db", "Ebb", "Fb"}, _
                                       {"C#", "D#", "E", "F#", "G#", "A", "B"}, _
                                       {"Db", "Eb", "Fb", "Gb", "Ab", "Bbb", "Cb"}, _
                                       {"Ab", "Bb", "Cb", "Db", "Eb", "Fb", "Gb"}, _
                                       {"Eb", "F", "Gb", "Ab", "Bb", "Cb", "Db"}, _
                                       {"Bb", "C", "Db", "Eb", "F", "Gb", "Ab"}, _
                                       {"F", "G", "Ab", "Bb", "C", "Db", "Eb"}, _
                                       {"G#", "A#", "B", "C#", "D#", "E", "F#"}, _
                                       {"Bbb", "Cb", "Dbb", "Ebb", "Fb", "Gbb", "Abb"}, _
                                       {"Abb", "Bbb", "Cbb", "D", "Ebb", "Fbb", "Gbb"}, _
                                       {"Fb", "Gb", "Abb", "Bbb", "Cb", "Dbb", "Ebb"}, _
                                       {"Ebb", "Fb", "Gbb", "Abb", "Bbb", "Cbb", "Dbb"}, _
                                       {"D#", "E#", "F#", "G#", "A#", "B", "C#"}, _
                                       {"E#", "F##", "G#", "A#", "B#", "C#", "D#"}, _
                                       {"A#", "B#", "C#", "D#", "E#", "F#", "G#"}, _
                                       {"B#", "C##", "D#", "E#", "F##", "G#", "A#"}, _
                                       {"F##", "G##", "A#", "B#", "C##", "D#", "E#"}, _
                                       {"Gbb", "Abb", "Bbbb", "Cbb", "Dbb", "Ebbb", "Fbb"}, _
                                       {"Dbb", "Ebb", "Fbb", "Gbb", "Abb", "Bbbb", "Cbb"}, _
                                       {"G##", "A##", "B#", "C##", "D##", "E#", "F##"}, _
                                       {"C##", "D##", "E#", "F##", "G##", "A#", "B#"}, _
                                       {"D##", "E##", "F##", "G##", "A##", "B#", "C##"}}


    Dim EscalaMenorHarmonica(,) As String = {{"C", "D", "Eb", "F", "G", "Ab", "B"}, _
                                       {"G", "A", "Bb", "C", "D", "Eb", "F#"}, _
                                       {"D", "E", "F", "G", "A", "Bb", "C#"}, _
                                       {"A", "B", "C", "D", "E", "F", "G#"}, _
                                       {"E", "F#", "G", "A", "B", "C", "D#"}, _
                                       {"B", "C#", "D", "E", "F#", "G", "A#"}, _
                                       {"Cb", "Db", "Ebb", "Fb", "Gb", "Abb", "Bb"}, _
                                       {"F#", "G#", "A", "B", "C#", "D", "E#"}, _
                                       {"Gb", "Ab", "Bbb", "Cb", "Db", "Ebb", "F"}, _
                                       {"C#", "D#", "E", "F#", "G#", "A", "B#"}, _
                                       {"Db", "Eb", "Fb", "Gb", "Ab", "Bbb", "C"}, _
                                       {"Ab", "Bb", "Cb", "Db", "Eb", "Fb", "G"}, _
                                       {"Eb", "F", "Gb", "Ab", "Bb", "Cb", "D"}, _
                                       {"Bb", "C", "Db", "Eb", "F", "Gb", "A"}, _
                                       {"F", "G", "Ab", "Bb", "C", "Db", "E"}, _
                                       {"G#", "A#", "B", "C#", "D#", "E", "F##"}, _
                                       {"Bbb", "Cb", "Dbb", "Ebb", "Fb", "Gbb", "Ab"}, _
                                       {"Abb", "Bbb", "Cbb", "D", "Ebb", "Fbb", "Gb"}, _
                                       {"Fb", "Gb", "Abb", "Bbb", "Cb", "Dbb", "Eb"}, _
                                       {"Ebb", "Fb", "Gbb", "Abb", "Bbb", "Cbb", "Db"}, _
                                       {"D#", "E#", "F#", "G#", "A#", "B", "C##"}, _
                                       {"E#", "F##", "G#", "A#", "B#", "C#", "D##"}, _
                                       {"A#", "B#", "C#", "D#", "E#", "F#", "G##"}, _
                                       {"B#", "C##", "D#", "E#", "F##", "G#", "A##"}, _
                                       {"F##", "G##", "A#", "B#", "C##", "D#", "E##"}, _
                                       {"Gbb", "Abb", "Bbbb", "Cbb", "Dbb", "Ebbb", "Fb"}, _
                                       {"Dbb", "Ebb", "Fbb", "Gbb", "Abb", "Bbbb", "Cb"}, _
                                       {"G##", "A##", "B#", "C##", "D##", "E#", "F###"}, _
                                       {"C##", "D##", "E#", "F##", "G##", "A#", "B##"}, _
                                       {"D##", "E##", "F##", "G##", "A##", "B#", "C###"}}


    Dim EscalaMenorMelodica(,) As String = {{"C", "D", "Eb", "F", "G", "A", "B"}, _
                                       {"G", "A", "Bb", "C", "D", "E", "F#"}, _
                                       {"D", "E", "F", "G", "A", "B", "C#"}, _
                                       {"A", "B", "C", "D", "E", "F#", "G#"}, _
                                       {"E", "F#", "G", "A", "B", "C#", "D#"}, _
                                       {"B", "C#", "D", "E", "F#", "G#", "A#"}, _
                                       {"Cb", "Db", "Ebb", "Fb", "Gb", "Ab", "Bb"}, _
                                       {"F#", "G#", "A", "B", "C#", "D#", "E#"}, _
                                       {"Gb", "Ab", "Bbb", "Cb", "Db", "Eb", "F"}, _
                                       {"C#", "D#", "E", "F#", "G#", "A#", "B#"}, _
                                       {"Db", "Eb", "Fb", "Gb", "Ab", "Bb", "C"}, _
                                       {"Ab", "Bb", "Cb", "Db", "Eb", "F", "G"}, _
                                       {"Eb", "F", "Gb", "Ab", "Bb", "C", "D"}, _
                                       {"Bb", "C", "Db", "Eb", "F", "G", "A"}, _
                                       {"F", "G", "Ab", "Bb", "C", "D", "E"}, _
                                       {"G#", "A#", "B", "C#", "D#", "E#", "F##"}, _
                                       {"Bbb", "Cb", "Dbb", "Ebb", "Fb", "Gb", "Ab"}, _
                                       {"Abb", "Bbb", "Cbb", "D", "Ebb", "Fb", "Gb"}, _
                                       {"Fb", "Gb", "Abb", "Bbb", "Cb", "Db", "Eb"}, _
                                       {"Ebb", "Fb", "Gbb", "Abb", "Bbb", "Cb", "Db"}, _
                                       {"D#", "E#", "F#", "G#", "A#", "B#", "C##"}, _
                                       {"E#", "F##", "G#", "A#", "B#", "C##", "D##"}, _
                                       {"A#", "B#", "C#", "D#", "E#", "F##", "G##"}, _
                                       {"B#", "C##", "D#", "E#", "F##", "G##", "A##"}, _
                                       {"F##", "G##", "A#", "B#", "C##", "D##", "E##"}, _
                                       {"Gbb", "Abb", "Bbbb", "Cbb", "Dbb", "Ebb", "Fb"}, _
                                       {"Dbb", "Ebb", "Fbb", "Gbb", "Abb", "Bbb", "Cb"}, _
                                       {"G##", "A##", "B#", "C##", "D##", "E##", "F###"}, _
                                       {"C##", "D##", "E#", "F##", "G##", "A##", "B##"}, _
                                       {"D##", "E##", "F##", "G##", "A##", "B##", "C###"}}

    Dim EscalaCoringa(29, 6)

    Dim GabaritoRespostas(6, 7) As String

    Private Sub SelecionarValores()

        Try

            DataGridView1.Rows.Item(6).Cells(0).Value = "1"
            GabaritoRespostas(6, 0) = "1"

            g = 0
            Do While g <= 5
                f = 100
                Do While f < 0 OrElse f > 6

                    ' Create a byte array to hold the random value.
                    Dim randomNumber(0) As Byte

                    ' Create a new instance of the RNGCryptoServiceProvider.
                    Dim Gen As New Security.Cryptography.RNGCryptoServiceProvider()

                    ' Fill the array with a random value.
                    Gen.GetBytes(randomNumber)

                    ' Convert the byte to an integer value to make the modulus operation easier.
                    f = Int(Convert.ToInt32(randomNumber(0)))

                Loop


                If DataGridView1.Rows.Item(f).Cells(0).Value Is Nothing Then
                    DataGridView1.Rows.Item(f).Cells(0).Value = g + 2
                    GabaritoRespostas(f, 0) = g + 2
                Else
                    g -= 1
                End If

                g += 1

            Loop



            h = 100
            Do While h < 1 OrElse h > 21

                ' Create a byte array to hold the random value.
                Dim randomNumber(0) As Byte

                ' Create a new instance of the RNGCryptoServiceProvider.
                Dim Gen As New Security.Cryptography.RNGCryptoServiceProvider()

                ' Fill the array with a random value.
                Gen.GetBytes(randomNumber)

                ' Convert the byte to an integer value to make the modulus operation easier.
                h = Int(Convert.ToInt32(randomNumber(0)))

                If OpçãoEscalaMenorNatural.Checked AndAlso (h = 17 OrElse h = 19 OrElse h = 20) Then h = 100

            Loop

            If h = 1 Then
                Nota = "C"
            ElseIf h = 2 Then
                Nota = "D"
            ElseIf h = 3 Then
                Nota = "E"
            ElseIf h = 4 Then
                Nota = "F"
            ElseIf h = 5 Then
                Nota = "G"
            ElseIf h = 6 Then
                Nota = "A"
            ElseIf h = 7 Then
                Nota = "B"
            ElseIf h = 8 Then
                Nota = "C#"
            ElseIf h = 9 Then
                Nota = "F#"
            ElseIf h = 10 Then
                Nota = "Cb"
            ElseIf h = 11 Then
                Nota = "Db"
            ElseIf h = 12 Then
                Nota = "Eb"
            ElseIf h = 13 Then
                Nota = "Gb"
            ElseIf h = 14 Then
                Nota = "Ab"
            ElseIf h = 15 Then
                Nota = "Bb"
            ElseIf h = 16 Then
                Nota = "D#"
            ElseIf h = 17 Then
                Nota = "E#"
            ElseIf h = 18 Then
                Nota = "G#"
            ElseIf h = 19 Then
                Nota = "A#"
            ElseIf h = 20 Then
                Nota = "B#"
            ElseIf h = 21 Then
                Nota = "Fb"
            End If


            GabaritoRespostas(6, 1) = Nota
            GabaritoRespostas(5, 2) = Nota
            GabaritoRespostas(4, 3) = Nota
            GabaritoRespostas(3, 4) = Nota
            GabaritoRespostas(2, 5) = Nota
            GabaritoRespostas(1, 6) = Nota
            GabaritoRespostas(0, 7) = Nota


            a = 1
            b = 6
            Do While a <= 7
                c = 0
                Do While GabaritoRespostas(6, 1) <> EscalaCoringa(c, GabaritoRespostas(b, 0) - 1) AndAlso c <= 29
                    c += 1
                Loop


                d = 0
                Do While d <= 6

                    If (OpçãoEscalaMaior.Checked AndAlso EscalaCoringa(c, 0) <> "G#" AndAlso EscalaCoringa(c, 0) <> "D#" AndAlso EscalaCoringa(c, 0) <> "A#" AndAlso _
                        EscalaCoringa(c, 0) <> "Bbb" AndAlso EscalaCoringa(c, 0) <> "Abb" AndAlso EscalaCoringa(c, 0) <> "Fb" AndAlso _
                        EscalaCoringa(c, 0) <> "Ebb" AndAlso EscalaCoringa(c, 0) <> "B#" AndAlso EscalaCoringa(c, 0) <> "E#" AndAlso _
                        EscalaCoringa(c, 0) <> "F##" AndAlso EscalaCoringa(c, 0) <> "Gbb" AndAlso EscalaCoringa(c, 0) <> "Dbb" AndAlso _
                        EscalaCoringa(c, 0) <> "G##" AndAlso EscalaCoringa(c, 0) <> "C##" AndAlso EscalaCoringa(c, 0) <> "D##") OrElse _
                    (Not OpçãoEscalaMaior.Checked AndAlso EscalaCoringa(c, 0) <> "Gb" AndAlso EscalaCoringa(c, 0) <> "Db" AndAlso EscalaCoringa(c, 0) <> "Cb" AndAlso _
                        EscalaCoringa(c, 0) <> "Bbb" AndAlso EscalaCoringa(c, 0) <> "Abb" AndAlso EscalaCoringa(c, 0) <> "Fb" AndAlso _
                        EscalaCoringa(c, 0) <> "Ebb" AndAlso EscalaCoringa(c, 0) <> "B#" AndAlso EscalaCoringa(c, 0) <> "E#" AndAlso _
                        EscalaCoringa(c, 0) <> "F##" AndAlso EscalaCoringa(c, 0) <> "Gbb" AndAlso EscalaCoringa(c, 0) <> "Dbb" AndAlso _
                        EscalaCoringa(c, 0) <> "G##" AndAlso EscalaCoringa(c, 0) <> "C##" AndAlso EscalaCoringa(c, 0) <> "D##") Then

                        If GabaritoRespostas(d, 0) = 1 Then
                            GabaritoRespostas(d, a) = EscalaCoringa(c, 0)
                        ElseIf GabaritoRespostas(d, 0) = 2 Then
                            GabaritoRespostas(d, a) = EscalaCoringa(c, 1)
                        ElseIf GabaritoRespostas(d, 0) = 3 Then
                            GabaritoRespostas(d, a) = EscalaCoringa(c, 2)
                        ElseIf GabaritoRespostas(d, 0) = 4 Then
                            GabaritoRespostas(d, a) = EscalaCoringa(c, 3)
                        ElseIf GabaritoRespostas(d, 0) = 5 Then
                            GabaritoRespostas(d, a) = EscalaCoringa(c, 4)
                        ElseIf GabaritoRespostas(d, 0) = 6 Then
                            GabaritoRespostas(d, a) = EscalaCoringa(c, 5)
                        ElseIf GabaritoRespostas(d, 0) = 7 Then
                            GabaritoRespostas(d, a) = EscalaCoringa(c, 6)
                        End If


                        If RadioButton2.Checked Then
                            If OpçãoEscalaMaior.Checked Then
                                If GabaritoRespostas(d, 0) = 2 OrElse GabaritoRespostas(d, 0) = 3 OrElse GabaritoRespostas(d, 0) = 6 Then
                                    GabaritoRespostas(d, a) = GabaritoRespostas(d, a) & "m"
                                ElseIf GabaritoRespostas(d, 0) = 5 Then
                                    GabaritoRespostas(d, a) = GabaritoRespostas(d, a) & "7"
                                ElseIf GabaritoRespostas(d, 0) = 7 Then
                                    GabaritoRespostas(d, a) = GabaritoRespostas(d, a) & "m7(b5)"
                                End If
                            ElseIf OpçãoEscalaMenorNatural.Checked Then
                                'por enquanto não vou utilizar
                            ElseIf OpçãoEscalaMenorHarmonica.Checked Then
                                'por enquanto não vou utilizar
                            ElseIf OpçãoEscalaMenorMelodica.Checked Then
                                'por enquanto não vou utilizar
                            End If
                        End If

                    End If

                    d += 1
                Loop

                a += 1
                b -= 1

            Loop


            DataGridView1.Rows.Item(6).Cells(1).Value = GabaritoRespostas(6, 1)
            DataGridView1.Rows.Item(5).Cells(2).Value = GabaritoRespostas(5, 2)
            DataGridView1.Rows.Item(4).Cells(3).Value = GabaritoRespostas(4, 3)
            DataGridView1.Rows.Item(3).Cells(4).Value = GabaritoRespostas(3, 4)
            DataGridView1.Rows.Item(2).Cells(5).Value = GabaritoRespostas(2, 5)
            DataGridView1.Rows.Item(1).Cells(6).Value = GabaritoRespostas(1, 6)
            DataGridView1.Rows.Item(0).Cells(7).Value = GabaritoRespostas(0, 7)


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting

        Try

            If e.CellStyle.BackColor <> Color.FromArgb(153, 204, 255) AndAlso e.ColumnIndex <> 0 Then
                e.Value = StrConv(e.Value, VbStrConv.ProperCase)
                e.Value = Replace(e.Value, "3", "#")
                e.Value = Replace(e.Value, "o", "m7(b5)")
                e.Value = Replace(e.Value, "O", "m7(b5)")
                e.Value = Replace(e.Value, "M", "m")
                e.Value = Replace(e.Value, "B5", "b5")
                e.Value = Replace(e.Value, " ", "")
            End If

            If e.RowIndex < 7 Then

                If e.Value IsNot Nothing Then

                    If e.Value.ToString.Length > 5 Then
                        e.CellStyle.Font = New Font("Arial", 10, FontStyle.Bold, GraphicsUnit.Point)
                    ElseIf e.Value.ToString.Length > 3 Then
                        e.CellStyle.Font = New Font("Arial", 18, FontStyle.Bold, GraphicsUnit.Point)
                    ElseIf e.Value.ToString.Length > 2 Then
                        e.CellStyle.Font = New Font("Arial", 20, FontStyle.Bold, GraphicsUnit.Point)
                    End If

                    If e.Value = "1" Then
                        e.CellStyle.BackColor = Color.FromArgb(0, 0, 0)
                        e.CellStyle.ForeColor = Color.FromArgb(255, 255, 255)
                    ElseIf e.Value = "2" OrElse e.Value = "3" OrElse e.Value = "4" OrElse e.Value = "5" OrElse e.Value = "6" OrElse e.Value = "7" Then
                        e.CellStyle.BackColor = Color.FromArgb(153, 204, 255)
                    ElseIf e.Value = GabaritoRespostas(e.RowIndex, e.ColumnIndex) Then
                        e.CellStyle.BackColor = Color.FromArgb(0, 255, 0)
                    Else
                        e.CellStyle.BackColor = Color.FromArgb(255, 0, 0)
                    End If

                Else
                    e.CellStyle.BackColor = Color.FromArgb(255, 255, 255)
                End If

                End If


            If (e.RowIndex = 6 AndAlso e.ColumnIndex = 1) OrElse (e.RowIndex = 5 AndAlso e.ColumnIndex = 2) OrElse (e.RowIndex = 4 AndAlso e.ColumnIndex = 3) _
                OrElse (e.RowIndex = 3 AndAlso e.ColumnIndex = 4) OrElse (e.RowIndex = 2 AndAlso e.ColumnIndex = 5) OrElse (e.RowIndex = 1 AndAlso e.ColumnIndex = 6) _
                OrElse (e.RowIndex = 0 AndAlso e.ColumnIndex = 7) Then
                e.CellStyle.BackColor = Color.FromArgb(255, 204, 153)
            End If

            If GabaritoRespostas(5, 1) = "" AndAlso (e.RowIndex = 6 AndAlso e.ColumnIndex = 1) Then e.CellStyle.BackColor = Color.FromArgb(100, 100, 100) : e.Value = ""
            If GabaritoRespostas(6, 2) = "" AndAlso (e.RowIndex = 5 AndAlso e.ColumnIndex = 2) Then e.CellStyle.BackColor = Color.FromArgb(100, 100, 100) : e.Value = ""
            If GabaritoRespostas(6, 3) = "" AndAlso (e.RowIndex = 4 AndAlso e.ColumnIndex = 3) Then e.CellStyle.BackColor = Color.FromArgb(100, 100, 100) : e.Value = ""
            If GabaritoRespostas(6, 4) = "" AndAlso (e.RowIndex = 3 AndAlso e.ColumnIndex = 4) Then e.CellStyle.BackColor = Color.FromArgb(100, 100, 100) : e.Value = ""
            If GabaritoRespostas(6, 5) = "" AndAlso (e.RowIndex = 2 AndAlso e.ColumnIndex = 5) Then e.CellStyle.BackColor = Color.FromArgb(100, 100, 100) : e.Value = ""
            If GabaritoRespostas(6, 6) = "" AndAlso (e.RowIndex = 1 AndAlso e.ColumnIndex = 6) Then e.CellStyle.BackColor = Color.FromArgb(100, 100, 100) : e.Value = ""
            If GabaritoRespostas(6, 7) = "" AndAlso (e.RowIndex = 0 AndAlso e.ColumnIndex = 7) Then e.CellStyle.BackColor = Color.FromArgb(100, 100, 100) : e.Value = ""


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub NovoExercicio(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click, OpçãoEscalaMaior.Click, OpçãoEscalaMenorNatural.Click, OpçãoEscalaMenorHarmonica.Click, OpçãoEscalaMenorMelodica.Click, RadioButton2.Click, RadioButton1.Click
        GerarNovoExercicio()
    End Sub

    Private Sub GerarNovoExercicio()
        Try
            DataGridView1.Rows.Clear()
            DataGridView1.Rows.Add(8)
            DataGridView1.Rows.Item(0).Height = 50
            DataGridView1.Rows.Item(1).Height = 50
            DataGridView1.Rows.Item(2).Height = 50
            DataGridView1.Rows.Item(3).Height = 50
            DataGridView1.Rows.Item(4).Height = 50
            DataGridView1.Rows.Item(5).Height = 50
            DataGridView1.Rows.Item(6).Height = 50
            DataGridView1.Rows.Item(7).Height = 50

            DataGridView1.Rows.Item(7).ReadOnly = True
            DataGridView1.Rows.Item(6).Cells(1).ReadOnly = True
            DataGridView1.Rows.Item(5).Cells(2).ReadOnly = True
            DataGridView1.Rows.Item(4).Cells(3).ReadOnly = True
            DataGridView1.Rows.Item(3).Cells(4).ReadOnly = True
            DataGridView1.Rows.Item(2).Cells(5).ReadOnly = True
            DataGridView1.Rows.Item(1).Cells(6).ReadOnly = True
            DataGridView1.Rows.Item(0).Cells(7).ReadOnly = True

            DataGridView1.Rows.Item(0).Cells(1).Selected = True

            DataGridView1.Rows.Item(7).DefaultCellStyle.BackColor = Color.FromArgb(153, 204, 255)
            DataGridView1.Rows.Item(7).Cells(0).Value = "Armadura"
            DataGridView1.Rows.Item(7).Cells(0).Style.Font = New Font("Arial", 10, FontStyle.Bold)

            Array.Clear(GabaritoRespostas, 0, GabaritoRespostas.Length)
            Array.Clear(EscalaCoringa, 0, EscalaCoringa.Length)

            GroupBox2.Visible = False
            If OpçãoEscalaMaior.Checked Then
                Array.Copy(EscalaMaior, EscalaCoringa, EscalaCoringa.Length)
                GroupBox2.Visible = True
            ElseIf OpçãoEscalaMenorNatural.Checked Then
                Array.Copy(EscalaMenorNatural, EscalaCoringa, EscalaCoringa.Length)
            ElseIf OpçãoEscalaMenorHarmonica.Checked Then
                Array.Copy(EscalaMenorHarmonica, EscalaCoringa, EscalaCoringa.Length)
            ElseIf OpçãoEscalaMenorMelodica.Checked Then
                Array.Copy(EscalaMenorMelodica, EscalaCoringa, EscalaCoringa.Length)
            End If


            SelecionarValores()

            AtualizaRegiões()


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub DataGridView1_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DataGridView1.KeyDown
        Try

            If e.KeyData = Keys.Delete AndAlso DataGridView1.CurrentCell IsNot Nothing Then
                If DataGridView1.CurrentCell.ReadOnly = False Then
                    DataGridView1.CurrentCell.Value = Nothing
                    AtualizaRegiões()
                End If
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub DataGridView1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles DataGridView1.Paint
        Try

            e.Graphics.DrawRectangle(Borda, 2, 2, DataGridView1.Width - 4, DataGridView1.Height - 4)
            e.Graphics.DrawRectangle(Borda, 2, 2, DataGridView1.Width - 4, DataGridView1.Height - 54)
            e.Graphics.DrawRectangle(Borda, 2, 2, 80, DataGridView1.Height - 4)


            ValorLeft = 96


            l = 1
            Do While l <= 7

                Nota2 = StrConv(DataGridView1.Rows.Item(6).Cells(l).Value, VbStrConv.ProperCase)
                Nota2 = Replace(Nota2, "3", "#")
                Nota2 = Replace(Nota2, " ", "")

                If DataGridView1.Rows.Item(6).Cells(l).Value <> "" AndAlso Nota2 = GabaritoRespostas(6, l) Then
                    If OpçãoEscalaMaior.Checked Then
                        If GabaritoRespostas(6, l) = "C" Then
                            e.Graphics.DrawImage(My.Resources.ClaveC1, ValorLeft, 364, 52, 33)
                            e.Graphics.DrawString("C - Am", Fonte, CorFonte, ValorLeft - 12, 353)
                        ElseIf GabaritoRespostas(6, l) = "D" Then
                            e.Graphics.DrawImage(My.Resources.ClaveD1, ValorLeft, 364, 52, 33)
                            e.Graphics.DrawString("D - Bm", Fonte, CorFonte, ValorLeft - 12, 353)
                        ElseIf GabaritoRespostas(6, l) = "E" Then
                            e.Graphics.DrawImage(My.Resources.ClaveE1, ValorLeft, 364, 52, 33)
                            e.Graphics.DrawString("E - C#m", Fonte, CorFonte, ValorLeft - 12, 353)
                        ElseIf GabaritoRespostas(6, l) = "F" Then
                            e.Graphics.DrawImage(My.Resources.ClaveF1, ValorLeft, 364, 52, 33)
                            e.Graphics.DrawString("F - Dm", Fonte, CorFonte, ValorLeft - 12, 353)
                        ElseIf GabaritoRespostas(6, l) = "G" Then
                            e.Graphics.DrawImage(My.Resources.ClaveG1, ValorLeft, 364, 52, 33)
                            e.Graphics.DrawString("G - Em", Fonte, CorFonte, ValorLeft - 12, 353)
                        ElseIf GabaritoRespostas(6, l) = "A" Then
                            e.Graphics.DrawImage(My.Resources.ClaveA1, ValorLeft, 364, 52, 33)
                            e.Graphics.DrawString("A - F#m", Fonte, CorFonte, ValorLeft - 12, 353)
                        ElseIf GabaritoRespostas(6, l) = "B" Then
                            e.Graphics.DrawImage(My.Resources.ClaveB1, ValorLeft, 364, 52, 33)
                            e.Graphics.DrawString("B - G#m", Fonte, CorFonte, ValorLeft - 12, 353)
                        ElseIf GabaritoRespostas(6, l) = "C#" Then
                            e.Graphics.DrawImage(My.Resources.ClaveCsus1, ValorLeft, 364, 52, 33)
                            e.Graphics.DrawString("C# - A#m", Fonte, CorFonte, ValorLeft - 12, 353)
                        ElseIf GabaritoRespostas(6, l) = "F#" Then
                            e.Graphics.DrawImage(My.Resources.ClaveFsus1, ValorLeft, 364, 52, 33)
                            e.Graphics.DrawString("F# - D#m", Fonte, CorFonte, ValorLeft - 12, 353)
                        ElseIf GabaritoRespostas(6, l) = "Cb" Then
                            e.Graphics.DrawImage(My.Resources.ClaveCbemol1, ValorLeft, 364, 52, 33)
                            e.Graphics.DrawString("Cb - Abm", Fonte, CorFonte, ValorLeft - 12, 353)
                        ElseIf GabaritoRespostas(6, l) = "Db" Then
                            e.Graphics.DrawImage(My.Resources.ClaveDbemol1, ValorLeft, 364, 52, 33)
                            e.Graphics.DrawString("Db - Bbm", Fonte, CorFonte, ValorLeft - 12, 353)
                        ElseIf GabaritoRespostas(6, l) = "Eb" Then
                            e.Graphics.DrawImage(My.Resources.ClaveEbemol1, ValorLeft, 364, 52, 33)
                            e.Graphics.DrawString("Eb - Cm", Fonte, CorFonte, ValorLeft - 12, 353)
                        ElseIf GabaritoRespostas(6, l) = "Gb" Then
                            e.Graphics.DrawImage(My.Resources.ClaveGbemol1, ValorLeft, 364, 52, 33)
                            e.Graphics.DrawString("Gb - Ebm", Fonte, CorFonte, ValorLeft - 12, 353)
                        ElseIf GabaritoRespostas(6, l) = "Ab" Then
                            e.Graphics.DrawImage(My.Resources.ClaveAbemol1, ValorLeft, 364, 52, 33)
                            e.Graphics.DrawString("Ab - Fm", Fonte, CorFonte, ValorLeft - 12, 353)
                        ElseIf GabaritoRespostas(6, l) = "Bb" Then
                            e.Graphics.DrawImage(My.Resources.ClaveBbemol1, ValorLeft, 364, 52, 33)
                            e.Graphics.DrawString("Bb - Gm", Fonte, CorFonte, ValorLeft - 12, 353)
                        End If
                    Else
                        If GabaritoRespostas(6, l) = "C" Then
                            e.Graphics.DrawImage(My.Resources.ClaveEbemol1, ValorLeft, 364, 52, 33)
                            e.Graphics.DrawString("Eb - Cm", Fonte, CorFonte, ValorLeft - 12, 353)
                        ElseIf GabaritoRespostas(6, l) = "D" Then
                            e.Graphics.DrawImage(My.Resources.ClaveF1, ValorLeft, 364, 52, 33)
                            e.Graphics.DrawString("F - Dm", Fonte, CorFonte, ValorLeft - 12, 353)
                        ElseIf GabaritoRespostas(6, l) = "E" Then
                            e.Graphics.DrawImage(My.Resources.ClaveG1, ValorLeft, 364, 52, 33)
                            e.Graphics.DrawString("G - Em", Fonte, CorFonte, ValorLeft - 12, 353)
                        ElseIf GabaritoRespostas(6, l) = "F" Then
                            e.Graphics.DrawImage(My.Resources.ClaveAbemol1, ValorLeft, 364, 52, 33)
                            e.Graphics.DrawString("Ab - Fm", Fonte, CorFonte, ValorLeft - 12, 353)
                        ElseIf GabaritoRespostas(6, l) = "G" Then
                            e.Graphics.DrawImage(My.Resources.ClaveBbemol1, ValorLeft, 364, 52, 33)
                            e.Graphics.DrawString("Bb - Gm", Fonte, CorFonte, ValorLeft - 12, 353)
                        ElseIf GabaritoRespostas(6, l) = "A" Then
                            e.Graphics.DrawImage(My.Resources.ClaveC1, ValorLeft, 364, 52, 33)
                            e.Graphics.DrawString("C - Am", Fonte, CorFonte, ValorLeft - 12, 353)
                        ElseIf GabaritoRespostas(6, l) = "B" Then
                            e.Graphics.DrawImage(My.Resources.ClaveD1, ValorLeft, 364, 52, 33)
                            e.Graphics.DrawString("D - Bm", Fonte, CorFonte, ValorLeft - 12, 353)
                        ElseIf GabaritoRespostas(6, l) = "C#" Then
                            e.Graphics.DrawImage(My.Resources.ClaveE1, ValorLeft, 364, 52, 33)
                            e.Graphics.DrawString("E - C#m", Fonte, CorFonte, ValorLeft - 12, 353)
                        ElseIf GabaritoRespostas(6, l) = "F#" Then
                            e.Graphics.DrawImage(My.Resources.ClaveA1, ValorLeft, 364, 52, 33)
                            e.Graphics.DrawString("A - F#m", Fonte, CorFonte, ValorLeft - 12, 353)
                            'ElseIf GabaritoRespostas(6, l) = "Cb" Then  NÃO POSSUI CLAVE
                            'e.Graphics.DrawImage(My.Resources.Clave, ValorLeft, 364, 52, 33)
                            'ElseIf GabaritoRespostas(6, l) = "Db" Then  NÃO POSSUI CLAVE
                            'e.Graphics.DrawImage(My.Resources.Clave, ValorLeft, 364, 52, 33)
                        ElseIf GabaritoRespostas(6, l) = "Eb" Then
                            e.Graphics.DrawImage(My.Resources.ClaveGbemol1, ValorLeft, 364, 52, 33)
                            e.Graphics.DrawString("Gb - Ebm", Fonte, CorFonte, ValorLeft - 12, 353)
                            'ElseIf GabaritoRespostas(6, l) = "Gb" Then  NÃO POSSUI CLAVE
                            'e.Graphics.DrawImage(My.Resources.Clave, ValorLeft, 364, 52, 33)
                        ElseIf GabaritoRespostas(6, l) = "Ab" Then
                            e.Graphics.DrawImage(My.Resources.ClaveCbemol1, ValorLeft, 364, 52, 33)
                            e.Graphics.DrawString("Cb - Abm", Fonte, CorFonte, ValorLeft - 12, 353)
                        ElseIf GabaritoRespostas(6, l) = "Bb" Then
                            e.Graphics.DrawImage(My.Resources.ClaveDbemol1, ValorLeft, 364, 52, 33)
                            e.Graphics.DrawString("Db - Bbm", Fonte, CorFonte, ValorLeft - 12, 353)
                        ElseIf GabaritoRespostas(6, l) = "G#" Then
                            e.Graphics.DrawImage(My.Resources.ClaveB1, ValorLeft, 364, 52, 33)
                            e.Graphics.DrawString("B - G#m", Fonte, CorFonte, ValorLeft - 12, 353)
                        ElseIf GabaritoRespostas(6, l) = "D#" Then
                            e.Graphics.DrawImage(My.Resources.ClaveFsus1, ValorLeft, 364, 52, 33)
                            e.Graphics.DrawString("F# - D#m", Fonte, CorFonte, ValorLeft - 12, 353)
                        ElseIf GabaritoRespostas(6, l) = "A#" Then
                            e.Graphics.DrawImage(My.Resources.ClaveCsus1, ValorLeft, 364, 52, 33)
                            e.Graphics.DrawString("C# - A#m", Fonte, CorFonte, ValorLeft - 12, 353)
                        End If
                    End If

                End If

                ValorLeft += 80
                l += 1
            Loop

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub DataGridView1_CellEndEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Try

            AtualizaRegiões()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Public Sub AtualizaRegiões()
        Try

            Dim Rect1 As New Rectangle(0, 350, 643, 50)
            DataGridView1.Invalidate(Rect1)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub SalvaSettings()
        Try

            My.Settings.NovoValorOpçãoTiposEscalas(0) = OpçãoEscalaMaior.Checked
            My.Settings.NovoValorOpçãoTiposEscalas(1) = OpçãoEscalaMenorNatural.Checked
            My.Settings.NovoValorOpçãoTiposEscalas(2) = OpçãoEscalaMenorHarmonica.Checked
            My.Settings.NovoValorOpçãoTiposEscalas(3) = OpçãoEscalaMenorMelodica.Checked
            My.Settings.NovoValorOpçãoTiposEscalas(4) = RadioButton1.Checked
            My.Settings.NovoValorOpçãoTiposEscalas(5) = RadioButton2.Checked

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub ExercicioParaMemorizarEscalas_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try

            OpçãoEscalaMaior.Checked = My.Settings.NovoValorOpçãoTiposEscalas(0)
            OpçãoEscalaMenorNatural.Checked = My.Settings.NovoValorOpçãoTiposEscalas(1)
            OpçãoEscalaMenorHarmonica.Checked = My.Settings.NovoValorOpçãoTiposEscalas(2)
            OpçãoEscalaMenorMelodica.Checked = My.Settings.NovoValorOpçãoTiposEscalas(3)
            RadioButton1.Checked = My.Settings.NovoValorOpçãoTiposEscalas(4)
            RadioButton2.Checked = My.Settings.NovoValorOpçãoTiposEscalas(5)

            GerarNovoExercicio()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub ExercicioParaMemorizarEscalas_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try

        Catch ex As Exception
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

            SalvaSettings()
        End Try
    End Sub
  
End Class