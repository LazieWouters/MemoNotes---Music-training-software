Option Strict On
Option Explicit On

Public Class Teclado
    Private DecimaSegundaThread As Thread
    Dim yyy, hh, spin, rr, OitavaEscala, NotaMidi As Integer
    Dim o As String
    Dim newFont As New Font("Arial", 6.5, FontStyle.Bold)
    Dim newFont2 As New Font("Arial", 12, FontStyle.Bold)
    Dim Img, ToolTipNota As Bitmap

    Private Sub TeclasPiano_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles P1.MouseHover, P3.MouseHover, P4.MouseHover, P6.MouseHover, P8.MouseHover, P9.MouseHover, P11.MouseHover, P13.MouseHover, P15.MouseHover, P16.MouseHover, P18.MouseHover, P20.MouseHover _
, P21.MouseHover, P23.MouseHover, P25.MouseHover, P27.MouseHover, P28.MouseHover, P30.MouseHover, P32.MouseHover, P33.MouseHover, P35.MouseHover, P37.MouseHover, P39.MouseHover, P40.MouseHover, P42.MouseHover, P44.MouseHover, P45.MouseHover, P47.MouseHover, P49.MouseHover, P51.MouseHover, P52.MouseHover, P54.MouseHover, P56.MouseHover, P57.MouseHover, P59.MouseHover, P61.MouseHover, P63.MouseHover, P64.MouseHover, P66.MouseHover, P68.MouseHover, P69.MouseHover, P71.MouseHover, P73.MouseHover, P75.MouseHover, P76.MouseHover _
, P78.MouseHover, P80.MouseHover, P81.MouseHover, P83.MouseHover, P85.MouseHover, P87.MouseHover, P88.MouseHover, P2.MouseHover, P5.MouseHover, P7.MouseHover, P10.MouseHover, P12.MouseHover, P14.MouseHover, P17.MouseHover, P19.MouseHover, P22.MouseHover, P24.MouseHover, P26.MouseHover, P29.MouseHover, P31.MouseHover, P34.MouseHover, P36.MouseHover, P38.MouseHover, P41.MouseHover, P43.MouseHover, P46.MouseHover, P48.MouseHover, P50.MouseHover, P53.MouseHover, P55.MouseHover, P58.MouseHover, P60.MouseHover _
, P62.MouseHover, P65.MouseHover, P67.MouseHover, P70.MouseHover, P72.MouseHover, P74.MouseHover, P77.MouseHover, P79.MouseHover, P82.MouseHover, P84.MouseHover, P86.MouseHover

        Try

            ' Obtém referência ao picturebox que invocou este método.

            pbox2 = DirectCast(sender, PictureBox)

            ToolTipNotas()
            ToolTipNota = ToolTipNota2

            'atualiza ToolTipNota
            Dim Rect8 As New Rectangle(479, 74, 104, 58)
            Me.Invalidate(Rect8)

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
, P62.MouseLeave, P65.MouseLeave, P67.MouseLeave, P70.MouseLeave, P72.MouseLeave, P74.MouseLeave, P77.MouseLeave, P79.MouseLeave, P82.MouseLeave, P84.MouseLeave, P86.MouseLeave

        Try

            ToolTipNota = Nothing
            'atualiza ToolTipNota
            Dim Rect8 As New Rectangle(479, 74, 104, 58)
            Me.Invalidate(Rect8)

            yyy = 0

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
, P62.MouseUp, P65.MouseUp, P67.MouseUp, P70.MouseUp, P72.MouseUp, P74.MouseUp, P77.MouseUp, P79.MouseUp, P82.MouseUp, P84.MouseUp, P86.MouseUp

        Try

            ' Obtém referência ao picturebox que invocou este método.

            Dim pbox As PictureBox = DirectCast(sender, PictureBox)
            If pbox.Width = 13 Then
                pbox.Image = My.Resources.Tecla_Preta
            Else
                pbox.Image = My.Resources.Tecla_Branca
            End If

            Try
                MidiPlayer.Play(New NoteOff(0, 1, CByte(CInt(pbox.Name.Replace("P", "")) + 20), CByte(NumericUpDown2.Value)))
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


            If ComboBox1.Text = "Piano Bösendorfer" Then
                mp3 = "Sons\Notas\Piano Bösendorfer\" & pbox.Name & ".mp3"
                'TocaSom()
            ElseIf ComboBox1.Text = "Strings (Trio)" Then
                mp3 = "Sons\Notas\Strings (Trio)\" & pbox.Name & ".mp3"
                ' TocaSom()
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

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub Teclado_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
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

    Public Sub SalvaSettings()
        Try

            My.Settings.NovoValorCheckBox100 = CheckBox100.Checked
            My.Settings.NovoValorCheckBox200 = CheckBox200.Checked
            My.Settings.NovoValorSimplificado = Simplificado.Checked
            My.Settings.NovoValorTimers(3) = Minutos.Text
            My.Settings.NovoValorTimers(4) = Segundos.Text
            My.Settings.NovoValorTimers(5) = Decimos.Text
            My.Settings.NovoValorInstrumentoMusical(1) = ComboBox1.Text

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Teclado_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            MidiPlayer.OpenMidi()
        Catch ex As Exception
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

            Tocanota = 1

            Minutos.Text = My.Settings.NovoValorTimers(3)
            Segundos.Text = My.Settings.NovoValorTimers(4)
            Decimos.Text = My.Settings.NovoValorTimers(5)

            Me.Top = Escalas.Top + 674
            Me.Left = Escalas.Left


            OitavaEscala = 123 + (119 * My.Settings.NovoValorOitavaEscala)

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

            ComboBox1.Text = My.Settings.NovoValorInstrumentoMusical(1)

            If ComboBox1.Text = "" Then ComboBox1.Text = "000 - Acoustic Grand Piano"
        End Try

    End Sub

    Private Sub PressionaTeclas()
        Try

            If (VGaa = 1 AndAlso VGz <= qtdeloop) OrElse (VGaa = 2 AndAlso VGz >= 0) Then
                DecimaSegundaThread = New Thread(AddressOf DecimaSegundaThreadCode)
                DecimaSegundaThread.Name = "DecimaSegunda Thread"
                DecimaSegundaThread.Start()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub DecimaSegundaThreadCode()
        Try

            Dim cControl() As Control = Me.Controls.Find("P" & rr, True)
            Dim pb As PictureBox = DirectCast(cControl(0), PictureBox)
            If rr = 4 OrElse rr = 9 OrElse rr = 16 OrElse rr = 21 OrElse rr = 28 OrElse rr = 33 OrElse rr = 40 OrElse rr = 45 OrElse rr = 52 OrElse rr = 57 OrElse rr = 64 OrElse rr = 69 OrElse rr = 76 OrElse rr = 81 Then
                pb.Image = My.Resources.Tecla_BrancaPressionada
            ElseIf rr = 6 OrElse rr = 11 OrElse rr = 13 OrElse rr = 18 OrElse rr = 23 OrElse rr = 25 OrElse rr = 30 OrElse rr = 35 OrElse rr = 37 OrElse rr = 42 OrElse rr = 47 OrElse rr = 49 OrElse rr = 54 OrElse rr = 59 OrElse rr = 61 OrElse rr = 66 OrElse rr = 71 OrElse rr = 73 OrElse rr = 78 OrElse rr = 83 OrElse rr = 85 Then
                pb.Image = My.Resources.Tecla_BrancaPressionada2
            ElseIf rr = 8 OrElse rr = 15 OrElse rr = 20 OrElse rr = 27 OrElse rr = 32 OrElse rr = 39 OrElse rr = 44 OrElse rr = 51 OrElse rr = 56 OrElse rr = 63 OrElse rr = 68 OrElse rr = 75 OrElse rr = 80 OrElse rr = 87 Then
                pb.Image = My.Resources.Tecla_BrancaPressionada3
            Else
                pb.Image = My.Resources.Tecla_PretaPressionada
            End If
            ValorP(0, rr) = 1

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Acorde_Fsusm()
        Try

            P34.Image = My.Resources.Tecla_PretaPressionada 'F# Gb
            P37.Image = My.Resources.Tecla_BrancaPressionada2 'A
            P41.Image = My.Resources.Tecla_PretaPressionada ' C#
            mp3 = "Sons\Acordes\Fsusm.mp3" 'TocaSom()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Acorde_A()
        Try

            P37.Image = My.Resources.Tecla_BrancaPressionada2 'A
            P41.Image = My.Resources.Tecla_PretaPressionada ' C#
            P44.Image = My.Resources.Tecla_BrancaPressionada3 'E
            mp3 = "Sons\Acordes\A.mp3" 'TocaSom()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Acorde_Csusº()
        Try

            P35.Image = My.Resources.Tecla_BrancaPressionada2 'G
            P41.Image = My.Resources.Tecla_PretaPressionada ' C#
            P44.Image = My.Resources.Tecla_BrancaPressionada3 'E
            mp3 = "Sons\Acordes\Csusº.mp3" 'TocaSom()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Acorde_Bm()
        Try

            P34.Image = My.Resources.Tecla_PretaPressionada 'F# Gb
            P39.Image = My.Resources.Tecla_BrancaPressionada3 'B
            P42.Image = My.Resources.Tecla_BrancaPressionada2 'D
            mp3 = "Sons\Acordes\Bm.mp3" 'TocaSom()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Acorde_D()
        Try

            P34.Image = My.Resources.Tecla_PretaPressionada 'F# Gb
            P37.Image = My.Resources.Tecla_BrancaPressionada2 'A
            P42.Image = My.Resources.Tecla_BrancaPressionada2 'D
            mp3 = "Sons\Acordes\D.mp3" 'TocaSom()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Acorde_Fsusº()
        Try

            P34.Image = My.Resources.Tecla_PretaPressionada 'F# Gb
            P37.Image = My.Resources.Tecla_BrancaPressionada2 'A
            P40.Image = My.Resources.Tecla_BrancaPressionada 'C
            mp3 = "Sons\Acordes\Fsusº.mp3" 'TocaSom()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Acorde_C()
        Try

            P35.Image = My.Resources.Tecla_BrancaPressionada2 'G
            P40.Image = My.Resources.Tecla_BrancaPressionada 'C
            P44.Image = My.Resources.Tecla_BrancaPressionada3 'E
            mp3 = "Sons\Acordes\C.mp3" 'TocaSom()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Acorde_Dm()
        Try

            P33.Image = My.Resources.Tecla_BrancaPressionada 'F
            P37.Image = My.Resources.Tecla_BrancaPressionada2 'A
            P42.Image = My.Resources.Tecla_BrancaPressionada2 'D
            mp3 = "Sons\Acordes\Dm.mp3" 'TocaSom()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Acorde_Em()
        Try

            P35.Image = My.Resources.Tecla_BrancaPressionada2 'G
            P39.Image = My.Resources.Tecla_BrancaPressionada3 'B
            P44.Image = My.Resources.Tecla_BrancaPressionada3 'E
            mp3 = "Sons\Acordes\Em.mp3" 'TocaSom()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Acorde_F()
        Try

            P37.Image = My.Resources.Tecla_BrancaPressionada2 'A
            P40.Image = My.Resources.Tecla_BrancaPressionada 'C
            P45.Image = My.Resources.Tecla_BrancaPressionada 'F
            mp3 = "Sons\Acordes\F.mp3" 'TocaSom()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Acorde_G()
        Try

            P35.Image = My.Resources.Tecla_BrancaPressionada2 'G
            P39.Image = My.Resources.Tecla_BrancaPressionada3 'B
            P42.Image = My.Resources.Tecla_BrancaPressionada2 'D
            mp3 = "Sons\Acordes\G.mp3" 'TocaSom()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Acorde_Am()
        Try

            P37.Image = My.Resources.Tecla_BrancaPressionada2 'A
            P40.Image = My.Resources.Tecla_BrancaPressionada 'C
            P44.Image = My.Resources.Tecla_BrancaPressionada3 'E
            mp3 = "Sons\Acordes\Am.mp3" 'TocaSom()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub Acorde_Bº()
        Try

            P39.Image = My.Resources.Tecla_BrancaPressionada3 'B
            P42.Image = My.Resources.Tecla_BrancaPressionada2 'D
            P45.Image = My.Resources.Tecla_BrancaPressionada 'F
            mp3 = "Sons\Acordes\Bº.mp3" 'TocaSom()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub CoresGrausEscala()

        Try

            Prog1A.BackColor = Color.Transparent
            Prog2A.BackColor = Color.Transparent
            Prog3A.BackColor = Color.Transparent
            Prog4A.BackColor = Color.Transparent
            Prog5A.BackColor = Color.Transparent
            Prog6A.BackColor = Color.Transparent
            Prog7A.BackColor = Color.Transparent
            Prog1.BackColor = Color.Transparent
            Prog2.BackColor = Color.Transparent
            Prog3.BackColor = Color.Transparent
            Prog4.BackColor = Color.Transparent
            Prog5.BackColor = Color.Transparent
            Prog6.BackColor = Color.Transparent
            Prog7.BackColor = Color.Transparent

            If a_I = 1 Then Prog1A.BackColor = Color.Khaki
            If a_II = 1 Then Prog2A.BackColor = Color.Khaki
            If a_III = 1 Then Prog3A.BackColor = Color.Khaki
            If a_IV = 1 Then Prog4A.BackColor = Color.Khaki
            If a_V = 1 Then Prog5A.BackColor = Color.Khaki
            If a_VI = 1 Then Prog6A.BackColor = Color.Khaki
            If a_VII = 1 Then Prog7A.BackColor = Color.Khaki

            Prog1.BackColor = Prog1A.BackColor
            Prog2.BackColor = Prog2A.BackColor
            Prog3.BackColor = Prog3A.BackColor
            Prog4.BackColor = Prog4A.BackColor
            Prog5.BackColor = Prog5A.BackColor
            Prog6.BackColor = Prog6A.BackColor
            Prog7.BackColor = Prog7A.BackColor

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick

        Try

            For i = 4 To 87
                If ValorP(0, i) = 1 Then ValorP(1, i) += 1
                If ValorP(1, i) = 20 Then
                    Dim cControl() As Control = Me.Controls.Find("P" & i, True)
                    Dim pb As PictureBox = DirectCast(cControl(0), PictureBox)
                    If pb.Width = 17 Then
                        pb.Image = My.Resources.Tecla_Branca
                    Else
                        pb.Image = My.Resources.Tecla_Preta
                    End If
                    ValorP(0, i) = 0 : ValorP(1, i) = 0
                End If
                VGzz += 1
            Next

            If VGzz = qtdeloop Then Timer2.Enabled = False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub GrausEscala()
        Try

            If VGccc = "Progressão 1" Then
                Prog1.Text = "I"
                Prog2.Text = "IIm"
                Prog3.Text = "IIIm"
                Prog4.Text = "IV"
                Prog5.Text = "V"
                Prog6.Text = "VIm"
                Prog7.Text = "VIIº"
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub LabelsTocaAcorde(ByVal sender As Object, ByVal e As System.EventArgs) Handles Prog1A.MouseDown, Prog2A.MouseDown, Prog3A.MouseDown, Prog4A.MouseDown, Prog5A.MouseDown, Prog6A.MouseDown, Prog7A.MouseDown

        Try

            ' Obtém referência ao picturebox que invocou este método.
            Dim pbox As Label = DirectCast(sender, Label)
            VGhhh = 1
            If pbox.Text = "C" Then
                Acorde_C()
            ElseIf pbox.Text = "F#m" Then
                Acorde_Fsusm()
            ElseIf pbox.Text = "A" Then
                Acorde_A()
            ElseIf pbox.Text = "C#º" Then
                Acorde_Csusº()
            ElseIf pbox.Text = "Bm" Then
                Acorde_Bm()
            ElseIf pbox.Text = "D" Then
                Acorde_D()
            ElseIf pbox.Text = "F#º" Then
                Acorde_Fsusº()
            ElseIf pbox.Text = "Dm" Then
                Acorde_Dm()
            ElseIf pbox.Text = "Em" Then
                Acorde_Em()
            ElseIf pbox.Text = "F" Then
                Acorde_F()
            ElseIf pbox.Text = "G" Then
                Acorde_G()
            ElseIf pbox.Text = "Am" Then
                Acorde_Am()
            ElseIf pbox.Text = "Bº" Then
                Acorde_Bº()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub Prog1A_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Prog7A.MouseUp, Prog6A.MouseUp, Prog5A.MouseUp, Prog4A.MouseUp, Prog3A.MouseUp, Prog2A.MouseUp, Prog1A.MouseUp
        DefineTecladoInicial()
    End Sub

    Private Sub Timer4_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer4.Tick

        Try

            DefineTecladoInicial()

            If VGccc = "Progressão 1" Then
                If VGab = 1 Then
                    a_I = 1
                ElseIf VGab = 2 Then
                    a_II = 1 : DefineTecladoInicial()
                ElseIf VGab = 3 Then
                    a_III = 1 : DefineTecladoInicial()
                ElseIf VGab = 4 Then
                    a_IV = 1 : DefineTecladoInicial()
                ElseIf VGab = 5 Then
                    a_V = 1 : DefineTecladoInicial()
                ElseIf VGab = 6 Then
                    a_VI = 1 : DefineTecladoInicial()
                ElseIf VGab = 7 Then
                    a_VII = 1 : DefineTecladoInicial()
                ElseIf VGab = 8 Then
                    CoresGrausEscala() : VGhhh = 0 : Timer4.Enabled = False
                End If
            End If

            If VGjjj = "Dó" Then
                If a_I = 1 Then
                    Acorde_C() : CoresGrausEscala() : a_I = 0
                ElseIf a_II = 1 Then
                    Acorde_Dm() : CoresGrausEscala() : a_II = 0
                ElseIf a_III = 1 Then
                    Acorde_Em() : CoresGrausEscala() : a_III = 0
                ElseIf a_IV = 1 Then
                    Acorde_F() : CoresGrausEscala() : a_IV = 0
                ElseIf a_V = 1 Then
                    Acorde_G() : CoresGrausEscala() : a_V = 0
                ElseIf a_VI = 1 Then
                    Acorde_Am() : CoresGrausEscala() : a_VI = 0
                ElseIf a_VII = 1 Then
                    Acorde_Bº() : CoresGrausEscala() : a_VII = 0
                End If
            ElseIf VGjjj = "Sol" Then
                If a_I = 1 Then
                    Acorde_G() : CoresGrausEscala() : a_I = 0
                ElseIf a_II = 1 Then
                    Acorde_Am() : CoresGrausEscala() : a_II = 0
                ElseIf a_III = 1 Then
                    Acorde_Bm() : CoresGrausEscala() : a_III = 0
                ElseIf a_IV = 1 Then
                    Acorde_C() : CoresGrausEscala() : a_IV = 0
                ElseIf a_V = 1 Then
                    Acorde_D() : CoresGrausEscala() : a_V = 0
                ElseIf a_VI = 1 Then
                    Acorde_Em() : CoresGrausEscala() : a_VI = 0
                ElseIf a_VII = 1 Then
                    Acorde_Fsusº() : CoresGrausEscala() : a_VII = 0
                End If
            ElseIf VGjjj = "Ré" Then
                If a_I = 1 Then
                    Acorde_D() : CoresGrausEscala() : a_I = 0
                ElseIf a_II = 1 Then
                    Acorde_Em() : CoresGrausEscala() : a_II = 0
                ElseIf a_III = 1 Then
                    Acorde_Fsusm() : CoresGrausEscala() : a_III = 0
                ElseIf a_IV = 1 Then
                    Acorde_G() : CoresGrausEscala() : a_IV = 0
                ElseIf a_V = 1 Then
                    Acorde_A() : CoresGrausEscala() : a_V = 0
                ElseIf a_VI = 1 Then
                    Acorde_Bm() : CoresGrausEscala() : a_VI = 0
                ElseIf a_VII = 1 Then
                    Acorde_Csusº() : CoresGrausEscala() : a_VII = 0
                End If
            End If


            VGab += 1
            Timer4.Interval = CInt(((CDbl(Minutos.Text) * 60) + CDbl(Segundos.Text) + (CDbl(Decimos.Text) / 100)) * 1000)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Public Sub DefineTecladoInicial() Handles Me.Load
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
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub PToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PToolStripMenuItem.Click
        Try

            VGccc = "Progressão 1"
            PictureBox5.Image = PToolStripMenuItem.Image
            GrausEscala()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub ab_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown, Prog7A.MouseDown, Prog7.MouseDown, Prog6A.MouseDown, Prog6.MouseDown, Prog5A.MouseDown, Prog5.MouseDown, Prog4A.MouseDown, Prog4.MouseDown, Prog3A.MouseDown, Prog3.MouseDown, Prog2A.MouseDown, Prog2.MouseDown, Prog1A.MouseDown, Prog1.MouseDown, PictureBox5.MouseDown, ListBox1.MouseDown, Simplificado.MouseDown, Segundos.MouseDown, PictureBox27.MouseDown, PictureBox26.MouseDown, Minutos.MouseDown, Menos.MouseDown, Mais.MouseDown, Decimos.MouseDown, CheckBox200.MouseDown, CheckBox100.MouseDown
        VGa = Me.MousePosition.X - Me.Location.X
        VGb = Me.MousePosition.Y - Me.Location.Y
    End Sub

    Private Sub ab_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove, Prog7A.MouseMove, Prog7.MouseMove, Prog6A.MouseMove, Prog6.MouseMove, Prog5A.MouseMove, Prog5.MouseMove, Prog4A.MouseMove, Prog4.MouseMove, Prog3A.MouseMove, Prog3.MouseMove, Prog2A.MouseMove, Prog2.MouseMove, Prog1A.MouseMove, Prog1.MouseMove, PictureBox5.MouseMove, ListBox1.MouseMove, Simplificado.MouseMove, Segundos.MouseMove, PictureBox27.MouseMove, PictureBox26.MouseMove, Minutos.MouseMove, Menos.MouseMove, Mais.MouseMove, Decimos.MouseMove, CheckBox200.MouseMove, CheckBox100.MouseMove

        If e.Button = MouseButtons.Left Then
            VGnewPoint = Me.MousePosition
            VGnewPoint.X -= VGa
            VGnewPoint.Y -= VGb
            Me.Location = VGnewPoint
            Escalas.Top = Me.Top - 674
            Escalas.Left = Me.Left
        End If

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Try

            rr = (VGx + (12 * My.Settings.NovoValorOitavaEscala)) + 3

            Try
                MidiPlayer.Play(New NoteOff(0, 1, CByte(NotaMidi), CByte(NumericUpDown2.Value)))
            Catch ex As Exception

                'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
            Finally
                'esta parte do bloco é executada independentemente se algum erro acontecer ou não

            End Try

            PressionaTeclas()

            If VGaa = 1 Then
                If VGz = qtdeloop + 1 Then
                    Timer1.Enabled = False
                ElseIf VGz = 1 Then
                    VGx += VGc
                ElseIf VGz = 2 Then
                    VGx += VGd
                ElseIf VGz = 3 Then
                    VGx += VGee
                ElseIf VGz = 4 Then
                    VGx += VGf
                ElseIf VGz = 5 Then
                    VGx += VGg
                ElseIf VGz = 6 Then
                    VGx += VGh
                ElseIf VGz = 7 Then
                    VGx += VGi
                ElseIf VGz = 8 Then
                    VGx += VGj
                ElseIf VGz = 9 Then
                    VGx += VGk
                ElseIf VGz = 10 Then
                    VGx += VGl
                ElseIf VGz = 11 Then
                    VGx += VGm
                ElseIf VGz = 12 Then
                    VGx += VGn
                ElseIf VGz = 13 Then
                    VGx += VGt
                End If

                TocaSomEscala()

                VGz += 1

            ElseIf VGaa = 2 Then
                If VGz = -1 Then
                    Timer1.Enabled = False
                ElseIf VGz = 1 Then
                    VGx -= VGc
                ElseIf VGz = 2 Then
                    VGx -= VGd
                ElseIf VGz = 3 Then
                    VGx -= VGee
                ElseIf VGz = 4 Then
                    VGx -= VGf
                ElseIf VGz = 5 Then
                    VGx -= VGg
                ElseIf VGz = 6 Then
                    VGx -= VGh
                ElseIf VGz = 7 Then
                    VGx -= VGi
                ElseIf VGz = 8 Then
                    VGx -= VGj
                ElseIf VGz = 9 Then
                    VGx -= VGk
                ElseIf VGz = 10 Then
                    VGx -= VGl
                ElseIf VGz = 11 Then
                    VGx -= VGm
                ElseIf VGz = 12 Then
                    VGx -= VGn
                ElseIf VGz = 13 Then
                    VGx -= VGt
                End If


                TocaSomEscala()

                VGz -= 1
            End If


            Timer1.Interval = CInt(((CDbl(Minutos.Text) * 60) + CDbl(Segundos.Text) + (CDbl(Decimos.Text) / 100)) * 1000) 'Velocidade.Value * 1000

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub TocaSomEscala()

        Try

            If (VGaa = 1 AndAlso VGz <= qtdeloop) OrElse (VGaa = 2 AndAlso VGz >= 0) Then
                If I_1_0 = 0 OrElse I_a1_1 = 1 OrElse I_b2_2 = 2 OrElse I_2_3 = 3 OrElse I_a2_4 = 4 OrElse I_b3_5 = 5 OrElse I_3_6 = 6 _
                   OrElse I_a3_7 = 7 OrElse I_b4_8 = 8 OrElse I_4_9 = 9 OrElse I_a4_10 = 10 OrElse I_b5_11 = 11 OrElse I_5_12 = 12 OrElse I_a5_13 = 13 _
                   OrElse I_b6_14 = 14 OrElse I_6_15 = 15 OrElse I_a6_16 = 16 OrElse I_b7_17 = 17 OrElse I_7_18 = 18 OrElse I_a7_19 = 19 _
                   OrElse I_b8_20 = 20 OrElse I_8_21 = 21 Then
                    If ComboBox1.Text = "Piano Bösendorfer" Then
                        mp3 = "Sons\Notas\Piano Bösendorfer\P" & rr & ".mp3"
                        'TocaSom()
                    ElseIf ComboBox1.Text = "Strings (Trio)" Then
                        mp3 = "Sons\Notas\Strings (Trio)\P" & rr & ".mp3"
                        'TocaSom()
                    Else
                        Try
                            MidiPlayer.Play(New ProgramChange(0, 1, CType(ComboBox1.Text.Substring(0, 3), GeneralMidiInstruments)))
                            MidiPlayer.Play(New NoteOn(0, 1, CByte(rr + 20), CByte(NumericUpDown2.Value)))
                        Catch ex As Exception

                            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
                        Finally
                            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

                        End Try
                        NotaMidi = CByte(rr + 20)
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

    Public Sub atualizaTeclado()
        Try

            For xTecla As Integer = 1 To 88
                Dim cControl() As Control = Me.Controls.Find("P" & xTecla.ToString, True)
                If cControl.Length > 0 Then
                    Dim pb As PictureBox = DirectCast(cControl(0), PictureBox)
                    pb.Refresh()
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub ListBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.Click
        ListBox1.Height = 250
        ListBox1.Width = 500
        ListBox1.BringToFront()
        ListBox1.Font = newFont2
    End Sub

    Private Sub ListBox1_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.MouseLeave
        ListBox1.Height = 55
        ListBox1.Width = 250
        ListBox1.Font = newFont
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal seder As System.Object, ByVal e As System.EventArgs) Handles CheckBox100.MouseUp, ToolStripMenuItem3.Click
        If VGoo <> "" Then
            Escalas.GeraEscalas()
        End If
        ToolStripMenuItem3.Checked = CheckBox100.Checked
    End Sub

    Private Sub CheckBox200_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox200.MouseUp, ToolStripMenuItem9.Click
        If VGoo <> "" Then
            Escalas.GeraEscalas()
        End If
        ToolStripMenuItem9.Checked = CheckBox200.Checked
    End Sub

    Private Sub Simplificado_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Simplificado.MouseUp, UsarCifrasSimplesExFGToolStripMenuItem.Click
        If VGoo <> "" Then
            Escalas.GeraEscalas()
        End If
        UsarCifrasSimplesExFGToolStripMenuItem.Checked = Simplificado.Checked
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
            Timer4.Interval = CInt(((CDbl(Minutos.Text) * 60) + CDbl(Segundos.Text) + (CDbl(Decimos.Text) / 100)) * 1000)

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
            Timer4.Interval = CInt(((CDbl(Minutos.Text) * 60) + CDbl(Segundos.Text) + (CDbl(Decimos.Text) / 100)) * 1000)

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
            If Minutos.Text = "00" AndAlso Segundos.Text = "00" AndAlso CDbl(Decimos.Text) < 10 Then Decimos.Text = "10"
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
            Timer4.Interval = CInt(((CDbl(Minutos.Text) * 60) + CDbl(Segundos.Text) + (CDbl(Decimos.Text) / 100)) * 1000)

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
            Timer4.Interval = CInt(((CDbl(Minutos.Text) * 60) + CDbl(Segundos.Text) + (CDbl(Decimos.Text) / 100)) * 1000)

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
                If Minutos.Text = "00" AndAlso Segundos.Text = "00" AndAlso Decimos.Text = "09" Then Decimos.Text = "10"
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
            Timer4.Interval = CInt(((CDbl(Minutos.Text) * 60) + CDbl(Segundos.Text) + (CDbl(Decimos.Text) / 100)) * 1000)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub Mais_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Mais.Click, ToolStripMenuItem8.Click
        Try

            If My.Settings.NovoValorOitavaEscala < 5 Then
                My.Settings.NovoValorOitavaEscala += 1
                OitavaEscala += 119
                Dim Rect As New Rectangle(88, 136, 880, 2)
                Me.Invalidate(Rect)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Menos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Menos.Click, ToolStripMenuItem7.Click
        Try

            If My.Settings.NovoValorOitavaEscala > 0 Then
                My.Settings.NovoValorOitavaEscala -= 1
                OitavaEscala -= 119
                Dim Rect As New Rectangle(88, 136, 880, 2)
                Me.Invalidate(Rect)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Fechar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem5.Click
        Try

            Me.Close()
            Escalas.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Minimizar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem4.Click

        Me.WindowState = CType(1, FormWindowState) '1 é para minimizar
        Escalas.WindowState = CType(1, FormWindowState) '1 é para minimizar

    End Sub

    Private Sub UsarCifrasSimplesExFGToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsarCifrasSimplesExFGToolStripMenuItem.Click
        Simplificado.Checked = Not Simplificado.Checked
        UsarCifrasSimplesExFGToolStripMenuItem.Checked = Simplificado.Checked
    End Sub

    Private Sub ToolStripMenuItem9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem9.Click
        CheckBox200.Checked = Not CheckBox200.Checked
        ToolStripMenuItem9.Checked = CheckBox200.Checked
    End Sub

    Private Sub ToolStripMenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem3.Click
        CheckBox100.Checked = Not CheckBox100.Checked
        ToolStripMenuItem3.Checked = CheckBox100.Checked
    End Sub

    Private Sub TesteToolStripMenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TesteToolStripMenuItem7.Click
        Try

            VGoo = ""
            ListBox1.Items.Clear()
            atualizaTeclado()
            VGjjj = ""
            NomeEscala = ""
            Intervalos = ""
            'zera as variáveis que somam os intervalos 
            VGc = 0 : VGd = 0 : VGee = 0 : VGf = 0 : VGg = 0 : VGh = 0 : VGi = 0 : VGj = 0 : VGk = 0 : VGl = 0 : VGm = 0 : VGn = 0 : VGt = 0

            Escalas.AnulaVariaveisDosIntervalos()
            Escalas.exibeControles()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Public Sub BolinhasColoridasTeclado(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles P1.Paint, P3.Paint, P4.Paint, P6.Paint, P8.Paint, P9.Paint, P11.Paint, P13.Paint, P15.Paint, P16.Paint, P18.Paint, P20.Paint _
, P21.Paint, P23.Paint, P25.Paint, P27.Paint, P28.Paint, P30.Paint, P32.Paint, P33.Paint, P35.Paint, P37.Paint, P39.Paint, P40.Paint, P42.Paint, P44.Paint, P45.Paint, P47.Paint, P49.Paint, P51.Paint, P52.Paint, P54.Paint, P56.Paint, P57.Paint, P59.Paint, P61.Paint, P63.Paint, P64.Paint, P66.Paint, P68.Paint, P69.Paint, P71.Paint, P73.Paint, P75.Paint, P76.Paint _
, P78.Paint, P80.Paint, P81.Paint, P83.Paint, P85.Paint, P87.Paint, P88.Paint, P2.Paint, P5.Paint, P7.Paint, P10.Paint, P12.Paint, P14.Paint, P17.Paint, P19.Paint, P22.Paint, P24.Paint, P26.Paint, P29.Paint, P31.Paint, P34.Paint, P36.Paint, P38.Paint, P41.Paint, P43.Paint, P46.Paint, P48.Paint, P50.Paint, P53.Paint, P55.Paint, P58.Paint, P60.Paint _
, P62.Paint, P65.Paint, P67.Paint, P70.Paint, P72.Paint, P74.Paint, P77.Paint, P79.Paint, P82.Paint, P84.Paint, P86.Paint

        Try

            Dim TônicaSemitransparente As SolidBrush = New SolidBrush(Color.FromArgb(50, 255, 0, 0))
            Dim FundoSemitransparente As SolidBrush = New SolidBrush(Color.FromArgb(100, 255, 255, 0))
            Dim Preto As SolidBrush = New SolidBrush(Escalas.ColorDialogC.Color)
            Dim Preto2 As SolidBrush = New SolidBrush(Color.FromArgb(2, 0, 0))
            Dim Verde As SolidBrush = New SolidBrush(Escalas.ColorDialogD.Color)
            Dim Azul As SolidBrush = New SolidBrush(Escalas.ColorDialogE.Color)
            Dim Vermelho As SolidBrush = New SolidBrush(Escalas.ColorDialogF.Color)
            Dim Vermelho2 As SolidBrush = New SolidBrush(Color.FromArgb(255, 0, 0))
            Dim Amarelo As SolidBrush = New SolidBrush(Escalas.ColorDialogG.Color)
            Dim Laranja As SolidBrush = New SolidBrush(Escalas.ColorDialogA.Color)
            Dim Rosa As SolidBrush = New SolidBrush(Escalas.ColorDialogB.Color)
            Dim Branco As SolidBrush = New SolidBrush(Color.FromArgb(255, 255, 250))


            Dim myFont As New Font("Arial", 11, FontStyle.Bold, GraphicsUnit.Pixel)
            Dim myFont2 As New Font("Arial", 7, FontStyle.Regular, GraphicsUnit.Pixel)

            ' Obtém referência ao picturebox que invocou este método.
            Dim pbox As PictureBox = DirectCast(sender, PictureBox)

            If VGoo <> "" Then
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias


                'If CheckBox1.checked Then
                'If p = "A" Then 'escala de Dó

                'C
                'Img = A1.Image
                ' If pbox.Name = "P4" OrElse  pbox.Name = "P16" OrElse  pbox.Name = "P28" OrElse  pbox.Name = "P40" OrElse  pbox.Name = "P52" OrElse  pbox.Name = "P64" OrElse  pbox.Name = "P76" OrElse  pbox.Name = "P88" Then
                'If Img.GetPixel(4, 39) = Color.FromArgb(255, 1, 0, 0) Then
                'e.Graphics.FillEllipse(Preto, 3, 72, 10, 10)
                ' Else
                ' e.Graphics.FillEllipse(Preto, 3, 70, 10, 10)
                ' End If
                ' End If

                'C# e Db
                'Img = A2.Image
                ' If pbox.Name = "P5" OrElse  pbox.Name = "P17" OrElse  pbox.Name = "P29" OrElse  pbox.Name = "P41" OrElse  pbox.Name = "P53" OrElse  pbox.Name = "P65" OrElse  pbox.Name = "P77" Then
                'If Img.GetPixel(3, 25) = Color.FromArgb(255, 1, 0, 0) Then
                'e.Graphics.FillEllipse(Branco, 1, 42, 10, 10)
                'e.Graphics.FillEllipse(Preto, 2, 43, 8, 8)
                ' Else
                ' e.Graphics.FillEllipse(Branco, 1, 39, 10, 10)
                ' e.Graphics.FillEllipse(Preto, 2, 40, 8, 8)
                ' End If
                'If Img.GetPixel(3, 25) = Color.FromArgb(255, 0, 99, 33) Then
                'e.Graphics.FillEllipse(Verde, 1, 42, 10, 10)
                'Else
                'e.Graphics.FillEllipse(Verde, 1, 39, 10, 10)
                ' End If
                ' End If


                'End If
                'End If

                Img = CType(pbox.Image, Bitmap)

                If VGoo = "Maior (T-T-s-T-T-T-s)" OrElse VGoo = "Jônio (T-T-s-T-T-T-s)" Then
                    If VGjjj = "Dó" Then
                        If pbox.Name = "P4" OrElse pbox.Name = "P16" OrElse pbox.Name = "P28" OrElse pbox.Name = "P40" OrElse pbox.Name = "P52" OrElse pbox.Name = "P64" OrElse pbox.Name = "P76" OrElse pbox.Name = "P88" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.FillRectangle(TônicaSemitransparente, 0, 0, 17, 85)
                                e.Graphics.DrawString("T", myFont, Vermelho2, 3, 60)
                                e.Graphics.FillEllipse(Preto, 3, 72, 10, 10)
                            Else
                                e.Graphics.FillRectangle(TônicaSemitransparente, 0, 0, 17, 85)
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("T", myFont, Vermelho2, 3, 58)
                                e.Graphics.FillEllipse(Preto, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P6" OrElse pbox.Name = "P18" OrElse pbox.Name = "P30" OrElse pbox.Name = "P42" OrElse pbox.Name = "P54" OrElse pbox.Name = "P66" OrElse pbox.Name = "P78" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("ST", myFont2, Preto2, 3, 65)
                                e.Graphics.FillEllipse(Verde, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("ST", myFont2, Preto2, 3, 63)
                                e.Graphics.FillEllipse(Verde, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P8" OrElse pbox.Name = "P20" OrElse pbox.Name = "P32" OrElse pbox.Name = "P44" OrElse pbox.Name = "P56" OrElse pbox.Name = "P68" OrElse pbox.Name = "P80" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("M", myFont2, Preto2, 4, 65)
                                e.Graphics.FillEllipse(Azul, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("M", myFont2, Preto2, 4, 63)
                                e.Graphics.FillEllipse(Azul, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P9" OrElse pbox.Name = "P21" OrElse pbox.Name = "P33" OrElse pbox.Name = "P45" OrElse pbox.Name = "P57" OrElse pbox.Name = "P69" OrElse pbox.Name = "P81" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("SD", myFont2, Preto2, 2, 65)
                                e.Graphics.FillEllipse(Vermelho, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("SD", myFont2, Preto2, 2, 63)
                                e.Graphics.FillEllipse(Vermelho, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P11" OrElse pbox.Name = "P23" OrElse pbox.Name = "P35" OrElse pbox.Name = "P47" OrElse pbox.Name = "P59" OrElse pbox.Name = "P71" OrElse pbox.Name = "P83" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("D", myFont2, Preto2, 5, 65)
                                e.Graphics.FillEllipse(Amarelo, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("D", myFont2, Preto2, 5, 63)
                                e.Graphics.FillEllipse(Amarelo, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P1" OrElse pbox.Name = "P13" OrElse pbox.Name = "P25" OrElse pbox.Name = "P37" OrElse pbox.Name = "P49" OrElse pbox.Name = "P61" OrElse pbox.Name = "P73" OrElse pbox.Name = "P85" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("SpD", myFont2, Preto2, 0, 65)
                                e.Graphics.FillEllipse(Laranja, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("SpD", myFont2, Preto2, 0, 63)
                                e.Graphics.FillEllipse(Laranja, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P3" OrElse pbox.Name = "P15" OrElse pbox.Name = "P27" OrElse pbox.Name = "P39" OrElse pbox.Name = "P51" OrElse pbox.Name = "P63" OrElse pbox.Name = "P75" OrElse pbox.Name = "P87" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("S", myFont2, Preto2, 5, 65)
                                e.Graphics.FillEllipse(Rosa, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("S", myFont2, Preto2, 5, 63)
                                e.Graphics.FillEllipse(Rosa, 3, 70, 10, 10)
                            End If
                        End If
                    ElseIf VGjjj = "Sol" Then
                        If pbox.Name = "P11" OrElse pbox.Name = "P23" OrElse pbox.Name = "P35" OrElse pbox.Name = "P47" OrElse pbox.Name = "P59" OrElse pbox.Name = "P71" OrElse pbox.Name = "P83" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.FillRectangle(TônicaSemitransparente, 0, 0, 17, 85)
                                e.Graphics.DrawString("T", myFont, Vermelho2, 3, 60)
                                e.Graphics.FillEllipse(Amarelo, 3, 72, 10, 10)
                            Else
                                e.Graphics.FillRectangle(TônicaSemitransparente, 0, 0, 17, 85)
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("T", myFont, Vermelho2, 3, 58)
                                e.Graphics.FillEllipse(Amarelo, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P1" OrElse pbox.Name = "P13" OrElse pbox.Name = "P25" OrElse pbox.Name = "P37" OrElse pbox.Name = "P49" OrElse pbox.Name = "P61" OrElse pbox.Name = "P73" OrElse pbox.Name = "P85" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("ST", myFont2, Preto2, 3, 65)
                                e.Graphics.FillEllipse(Laranja, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("ST", myFont2, Preto2, 3, 63)
                                e.Graphics.FillEllipse(Laranja, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P3" OrElse pbox.Name = "P15" OrElse pbox.Name = "P27" OrElse pbox.Name = "P39" OrElse pbox.Name = "P51" OrElse pbox.Name = "P63" OrElse pbox.Name = "P75" OrElse pbox.Name = "P87" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("M", myFont2, Preto2, 4, 65)
                                e.Graphics.FillEllipse(Rosa, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("M", myFont2, Preto2, 4, 63)
                                e.Graphics.FillEllipse(Rosa, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P4" OrElse pbox.Name = "P16" OrElse pbox.Name = "P28" OrElse pbox.Name = "P40" OrElse pbox.Name = "P52" OrElse pbox.Name = "P64" OrElse pbox.Name = "P76" OrElse pbox.Name = "P88" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("SD", myFont2, Preto2, 2, 65)
                                e.Graphics.FillEllipse(Preto, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("SD", myFont2, Preto2, 2, 63)
                                e.Graphics.FillEllipse(Preto, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P6" OrElse pbox.Name = "P18" OrElse pbox.Name = "P30" OrElse pbox.Name = "P42" OrElse pbox.Name = "P54" OrElse pbox.Name = "P66" OrElse pbox.Name = "P78" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("D", myFont2, Preto2, 5, 65)
                                e.Graphics.FillEllipse(Verde, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("D", myFont2, Preto2, 5, 63)
                                e.Graphics.FillEllipse(Verde, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P8" OrElse pbox.Name = "P20" OrElse pbox.Name = "P32" OrElse pbox.Name = "P44" OrElse pbox.Name = "P56" OrElse pbox.Name = "P68" OrElse pbox.Name = "P80" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("SpD", myFont2, Preto2, 0, 65)
                                e.Graphics.FillEllipse(Azul, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("SpD", myFont2, Preto2, 0, 63)
                                e.Graphics.FillEllipse(Azul, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P10" OrElse pbox.Name = "P22" OrElse pbox.Name = "P34" OrElse pbox.Name = "P46" OrElse pbox.Name = "P58" OrElse pbox.Name = "P70" OrElse pbox.Name = "P82" Then
                            If Img.GetPixel(1, 55) = Color.FromArgb(255, 1, 0, 0) Then
                                e.Graphics.DrawString("S", myFont2, Branco, 3, 35)
                                e.Graphics.FillEllipse(Vermelho, 1, 42, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 51)
                                e.Graphics.DrawString("S", myFont2, Branco, 3, 32)
                                e.Graphics.FillEllipse(Vermelho, 1, 39, 10, 10)
                            End If
                        End If
                    ElseIf VGjjj = "Ré" Then
                        If pbox.Name = "P6" OrElse pbox.Name = "P18" OrElse pbox.Name = "P30" OrElse pbox.Name = "P42" OrElse pbox.Name = "P54" OrElse pbox.Name = "P66" OrElse pbox.Name = "P78" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.FillRectangle(TônicaSemitransparente, 0, 0, 17, 85)
                                e.Graphics.DrawString("T", myFont, Vermelho2, 3, 60)
                                e.Graphics.FillEllipse(Verde, 3, 72, 10, 10)
                            Else
                                e.Graphics.FillRectangle(TônicaSemitransparente, 0, 0, 17, 85)
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("T", myFont, Vermelho2, 3, 58)
                                e.Graphics.FillEllipse(Verde, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P8" OrElse pbox.Name = "P20" OrElse pbox.Name = "P32" OrElse pbox.Name = "P44" OrElse pbox.Name = "P56" OrElse pbox.Name = "P68" OrElse pbox.Name = "P80" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("ST", myFont2, Preto2, 3, 65)
                                e.Graphics.FillEllipse(Azul, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("ST", myFont2, Preto2, 3, 63)
                                e.Graphics.FillEllipse(Azul, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P10" OrElse pbox.Name = "P22" OrElse pbox.Name = "P34" OrElse pbox.Name = "P46" OrElse pbox.Name = "P58" OrElse pbox.Name = "P70" OrElse pbox.Name = "P82" Then
                            If Img.GetPixel(1, 55) = Color.FromArgb(255, 1, 0, 0) Then
                                e.Graphics.DrawString("M", myFont2, Branco, 2, 35)
                                e.Graphics.FillEllipse(Vermelho, 1, 42, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 51)
                                e.Graphics.DrawString("M", myFont2, Branco, 2, 32)
                                e.Graphics.FillEllipse(Vermelho, 1, 39, 10, 10)
                            End If
                        ElseIf pbox.Name = "P11" OrElse pbox.Name = "P23" OrElse pbox.Name = "P35" OrElse pbox.Name = "P47" OrElse pbox.Name = "P59" OrElse pbox.Name = "P71" OrElse pbox.Name = "P83" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("SD", myFont2, Preto2, 2, 65)
                                e.Graphics.FillEllipse(Amarelo, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("SD", myFont2, Preto2, 2, 63)
                                e.Graphics.FillEllipse(Amarelo, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P1" OrElse pbox.Name = "P13" OrElse pbox.Name = "P25" OrElse pbox.Name = "P37" OrElse pbox.Name = "P49" OrElse pbox.Name = "P61" OrElse pbox.Name = "P73" OrElse pbox.Name = "P85" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("D", myFont2, Preto2, 5, 65)
                                e.Graphics.FillEllipse(Laranja, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("D", myFont2, Preto2, 5, 63)
                                e.Graphics.FillEllipse(Laranja, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P3" OrElse pbox.Name = "P15" OrElse pbox.Name = "P27" OrElse pbox.Name = "P39" OrElse pbox.Name = "P51" OrElse pbox.Name = "P63" OrElse pbox.Name = "P75" OrElse pbox.Name = "P87" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("SpD", myFont2, Preto2, 0, 65)
                                e.Graphics.FillEllipse(Rosa, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("SpD", myFont2, Preto2, 0, 63)
                                e.Graphics.FillEllipse(Rosa, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P5" OrElse pbox.Name = "P17" OrElse pbox.Name = "P29" OrElse pbox.Name = "P41" OrElse pbox.Name = "P53" OrElse pbox.Name = "P65" OrElse pbox.Name = "P77" OrElse pbox.Name = "T52A" Then
                            If Img.GetPixel(1, 55) = Color.FromArgb(255, 1, 0, 0) Then
                                e.Graphics.DrawString("S", myFont2, Branco, 3, 35)
                                e.Graphics.FillEllipse(Branco, 1, 42, 10, 10)
                                e.Graphics.FillEllipse(Preto, 2, 43, 8, 8)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 51)
                                e.Graphics.DrawString("S", myFont2, Branco, 3, 32)
                                e.Graphics.FillEllipse(Branco, 1, 39, 10, 10)
                                e.Graphics.FillEllipse(Preto, 2, 40, 8, 8)
                            End If
                        End If
                    ElseIf VGjjj = "Lá" Then
                        If pbox.Name = "P1" OrElse pbox.Name = "P13" OrElse pbox.Name = "P25" OrElse pbox.Name = "P37" OrElse pbox.Name = "P49" OrElse pbox.Name = "P61" OrElse pbox.Name = "P73" OrElse pbox.Name = "P85" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.FillRectangle(TônicaSemitransparente, 0, 0, 17, 85)
                                e.Graphics.DrawString("T", myFont, Vermelho2, 3, 60)
                                e.Graphics.FillEllipse(Laranja, 3, 72, 10, 10)
                            Else
                                e.Graphics.FillRectangle(TônicaSemitransparente, 0, 0, 17, 85)
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("T", myFont, Vermelho2, 3, 58)
                                e.Graphics.FillEllipse(Laranja, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P3" OrElse pbox.Name = "P15" OrElse pbox.Name = "P27" OrElse pbox.Name = "P39" OrElse pbox.Name = "P51" OrElse pbox.Name = "P63" OrElse pbox.Name = "P75" OrElse pbox.Name = "P87" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("ST", myFont2, Preto2, 3, 65)
                                e.Graphics.FillEllipse(Rosa, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("ST", myFont2, Preto2, 3, 63)
                                e.Graphics.FillEllipse(Rosa, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P5" OrElse pbox.Name = "P17" OrElse pbox.Name = "P29" OrElse pbox.Name = "P41" OrElse pbox.Name = "P53" OrElse pbox.Name = "P65" OrElse pbox.Name = "P77" OrElse pbox.Name = "T52A" Then
                            If Img.GetPixel(1, 55) = Color.FromArgb(255, 1, 0, 0) Then
                                e.Graphics.DrawString("M", myFont2, Branco, 2, 35)
                                e.Graphics.FillEllipse(Branco, 1, 42, 10, 10)
                                e.Graphics.FillEllipse(Preto, 2, 43, 8, 8)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 51)
                                e.Graphics.DrawString("M", myFont2, Branco, 2, 32)
                                e.Graphics.FillEllipse(Branco, 1, 39, 10, 10)
                                e.Graphics.FillEllipse(Preto, 2, 40, 8, 8)
                            End If
                        ElseIf pbox.Name = "P6" OrElse pbox.Name = "P18" OrElse pbox.Name = "P30" OrElse pbox.Name = "P42" OrElse pbox.Name = "P54" OrElse pbox.Name = "P66" OrElse pbox.Name = "P78" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("SD", myFont2, Preto2, 2, 65)
                                e.Graphics.FillEllipse(Verde, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("SD", myFont2, Preto2, 2, 63)
                                e.Graphics.FillEllipse(Verde, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P8" OrElse pbox.Name = "P20" OrElse pbox.Name = "P32" OrElse pbox.Name = "P44" OrElse pbox.Name = "P56" OrElse pbox.Name = "P68" OrElse pbox.Name = "P80" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("D", myFont2, Preto2, 5, 65)
                                e.Graphics.FillEllipse(Azul, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("D", myFont2, Preto2, 5, 63)
                                e.Graphics.FillEllipse(Azul, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P10" OrElse pbox.Name = "P22" OrElse pbox.Name = "P34" OrElse pbox.Name = "P46" OrElse pbox.Name = "P58" OrElse pbox.Name = "P70" OrElse pbox.Name = "P82" Then
                            If Img.GetPixel(1, 55) = Color.FromArgb(255, 1, 0, 0) Then
                                e.Graphics.DrawString("SpD", myFont2, Branco, -1, 35)
                                e.Graphics.FillEllipse(Vermelho, 1, 42, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 51)
                                e.Graphics.DrawString("SpD", myFont2, Branco, -1, 32)
                                e.Graphics.FillEllipse(Vermelho, 1, 39, 10, 10)
                            End If
                        ElseIf pbox.Name = "P12" OrElse pbox.Name = "P24" OrElse pbox.Name = "P36" OrElse pbox.Name = "P48" OrElse pbox.Name = "P60" OrElse pbox.Name = "P72" OrElse pbox.Name = "P84" Then
                            If Img.GetPixel(1, 55) = Color.FromArgb(255, 1, 0, 0) Then
                                e.Graphics.DrawString("S", myFont2, Branco, 3, 35)
                                e.Graphics.FillEllipse(Amarelo, 1, 42, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 51)
                                e.Graphics.DrawString("S", myFont2, Branco, 3, 32)
                                e.Graphics.FillEllipse(Amarelo, 1, 39, 10, 10)
                            End If
                        End If
                    ElseIf VGjjj = "Mi" Then
                        If pbox.Name = "P8" OrElse pbox.Name = "P20" OrElse pbox.Name = "P32" OrElse pbox.Name = "P44" OrElse pbox.Name = "P56" OrElse pbox.Name = "P68" OrElse pbox.Name = "P80" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.FillRectangle(TônicaSemitransparente, 0, 0, 17, 85)
                                e.Graphics.DrawString("T", myFont, Vermelho2, 3, 60)
                                e.Graphics.FillEllipse(Azul, 3, 72, 10, 10)
                            Else
                                e.Graphics.FillRectangle(TônicaSemitransparente, 0, 0, 17, 85)
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("T", myFont, Vermelho2, 3, 58)
                                e.Graphics.FillEllipse(Azul, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P10" OrElse pbox.Name = "P22" OrElse pbox.Name = "P34" OrElse pbox.Name = "P46" OrElse pbox.Name = "P58" OrElse pbox.Name = "P70" OrElse pbox.Name = "P82" Then
                            If Img.GetPixel(1, 55) = Color.FromArgb(255, 1, 0, 0) Then
                                e.Graphics.DrawString("ST", myFont2, Branco, 1, 35)
                                e.Graphics.FillEllipse(Vermelho, 1, 42, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 51)
                                e.Graphics.DrawString("ST", myFont2, Branco, 1, 32)
                                e.Graphics.FillEllipse(Vermelho, 1, 39, 10, 10)
                            End If
                        ElseIf pbox.Name = "P12" OrElse pbox.Name = "P24" OrElse pbox.Name = "P36" OrElse pbox.Name = "P48" OrElse pbox.Name = "P60" OrElse pbox.Name = "P72" OrElse pbox.Name = "P84" Then
                            If Img.GetPixel(1, 55) = Color.FromArgb(255, 1, 0, 0) Then
                                e.Graphics.DrawString("M", myFont2, Branco, 2, 35)
                                e.Graphics.FillEllipse(Amarelo, 1, 42, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 51)
                                e.Graphics.DrawString("M", myFont2, Branco, 2, 32)
                                e.Graphics.FillEllipse(Amarelo, 1, 39, 10, 10)
                            End If
                        ElseIf pbox.Name = "P1" OrElse pbox.Name = "P13" OrElse pbox.Name = "P25" OrElse pbox.Name = "P37" OrElse pbox.Name = "P49" OrElse pbox.Name = "P61" OrElse pbox.Name = "P73" OrElse pbox.Name = "P85" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("SD", myFont2, Preto2, 2, 65)
                                e.Graphics.FillEllipse(Laranja, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("SD", myFont2, Preto2, 2, 63)
                                e.Graphics.FillEllipse(Laranja, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P3" OrElse pbox.Name = "P15" OrElse pbox.Name = "P27" OrElse pbox.Name = "P39" OrElse pbox.Name = "P51" OrElse pbox.Name = "P63" OrElse pbox.Name = "P75" OrElse pbox.Name = "P87" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("D", myFont2, Preto2, 5, 65)
                                e.Graphics.FillEllipse(Rosa, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("D", myFont2, Preto2, 5, 63)
                                e.Graphics.FillEllipse(Rosa, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P5" OrElse pbox.Name = "P17" OrElse pbox.Name = "P29" OrElse pbox.Name = "P41" OrElse pbox.Name = "P53" OrElse pbox.Name = "P65" OrElse pbox.Name = "P77" OrElse pbox.Name = "T52A" Then
                            If Img.GetPixel(1, 55) = Color.FromArgb(255, 1, 0, 0) Then
                                e.Graphics.DrawString("SpD", myFont2, Branco, -1, 35)
                                e.Graphics.FillEllipse(Branco, 1, 42, 10, 10)
                                e.Graphics.FillEllipse(Preto, 2, 43, 8, 8)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 51)
                                e.Graphics.DrawString("SpD", myFont2, Branco, -1, 32)
                                e.Graphics.FillEllipse(Branco, 1, 39, 10, 10)
                                e.Graphics.FillEllipse(Preto, 2, 40, 8, 8)
                            End If
                        ElseIf pbox.Name = "P7" OrElse pbox.Name = "P19" OrElse pbox.Name = "P31" OrElse pbox.Name = "P43" OrElse pbox.Name = "P55" OrElse pbox.Name = "P67" OrElse pbox.Name = "P79" Then
                            If Img.GetPixel(1, 55) = Color.FromArgb(255, 1, 0, 0) Then
                                e.Graphics.DrawString("S", myFont2, Branco, 3, 35)
                                e.Graphics.FillEllipse(Verde, 1, 42, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 51)
                                e.Graphics.DrawString("S", myFont2, Branco, 3, 32)
                                e.Graphics.FillEllipse(Verde, 1, 39, 10, 10)
                            End If
                        End If
                    ElseIf VGjjj = "Si" Then
                        If pbox.Name = "P3" OrElse pbox.Name = "P15" OrElse pbox.Name = "P27" OrElse pbox.Name = "P39" OrElse pbox.Name = "P51" OrElse pbox.Name = "P63" OrElse pbox.Name = "P75" OrElse pbox.Name = "P87" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.FillRectangle(TônicaSemitransparente, 0, 0, 17, 85)
                                e.Graphics.DrawString("T", myFont, Vermelho2, 3, 60)
                                e.Graphics.FillEllipse(Rosa, 3, 72, 10, 10)
                            Else
                                e.Graphics.FillRectangle(TônicaSemitransparente, 0, 0, 17, 85)
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("T", myFont, Vermelho2, 3, 58)
                                e.Graphics.FillEllipse(Rosa, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P5" OrElse pbox.Name = "P17" OrElse pbox.Name = "P29" OrElse pbox.Name = "P41" OrElse pbox.Name = "P53" OrElse pbox.Name = "P65" OrElse pbox.Name = "P77" OrElse pbox.Name = "T52A" Then
                            If Img.GetPixel(1, 55) = Color.FromArgb(255, 1, 0, 0) Then
                                e.Graphics.DrawString("ST", myFont2, Branco, 1, 35)
                                e.Graphics.FillEllipse(Branco, 1, 42, 10, 10)
                                e.Graphics.FillEllipse(Preto, 2, 43, 8, 8)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 51)
                                e.Graphics.DrawString("ST", myFont2, Branco, 1, 32)
                                e.Graphics.FillEllipse(Branco, 1, 39, 10, 10)
                                e.Graphics.FillEllipse(Preto, 2, 40, 8, 8)
                            End If
                        ElseIf pbox.Name = "P7" OrElse pbox.Name = "P19" OrElse pbox.Name = "P31" OrElse pbox.Name = "P43" OrElse pbox.Name = "P55" OrElse pbox.Name = "P67" OrElse pbox.Name = "P79" Then
                            If Img.GetPixel(1, 55) = Color.FromArgb(255, 1, 0, 0) Then
                                e.Graphics.DrawString("M", myFont2, Branco, 2, 35)
                                e.Graphics.FillEllipse(Verde, 1, 42, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 51)
                                e.Graphics.DrawString("M", myFont2, Branco, 2, 32)
                                e.Graphics.FillEllipse(Verde, 1, 39, 10, 10)
                            End If
                        ElseIf pbox.Name = "P8" OrElse pbox.Name = "P20" OrElse pbox.Name = "P32" OrElse pbox.Name = "P44" OrElse pbox.Name = "P56" OrElse pbox.Name = "P68" OrElse pbox.Name = "P80" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("SD", myFont2, Preto2, 2, 65)
                                e.Graphics.FillEllipse(Azul, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("SD", myFont2, Preto2, 2, 63)
                                e.Graphics.FillEllipse(Azul, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P10" OrElse pbox.Name = "P22" OrElse pbox.Name = "P34" OrElse pbox.Name = "P46" OrElse pbox.Name = "P58" OrElse pbox.Name = "P70" OrElse pbox.Name = "P82" Then
                            If Img.GetPixel(1, 55) = Color.FromArgb(255, 1, 0, 0) Then
                                e.Graphics.DrawString("D", myFont2, Branco, 3, 35)
                                e.Graphics.FillEllipse(Vermelho, 1, 42, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 51)
                                e.Graphics.DrawString("D", myFont2, Branco, 3, 32)
                                e.Graphics.FillEllipse(Vermelho, 1, 39, 10, 10)
                            End If
                        ElseIf pbox.Name = "P12" OrElse pbox.Name = "P24" OrElse pbox.Name = "P36" OrElse pbox.Name = "P48" OrElse pbox.Name = "P60" OrElse pbox.Name = "P72" OrElse pbox.Name = "P84" Then
                            If Img.GetPixel(1, 55) = Color.FromArgb(255, 1, 0, 0) Then
                                e.Graphics.DrawString("SpD", myFont2, Branco, -1, 35)
                                e.Graphics.FillEllipse(Amarelo, 1, 42, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 51)
                                e.Graphics.DrawString("SpD", myFont2, Branco, -1, 32)
                                e.Graphics.FillEllipse(Amarelo, 1, 39, 10, 10)
                            End If
                        ElseIf pbox.Name = "P2" OrElse pbox.Name = "P14" OrElse pbox.Name = "P26" OrElse pbox.Name = "P38" OrElse pbox.Name = "P50" OrElse pbox.Name = "P62" OrElse pbox.Name = "P74" OrElse pbox.Name = "P86" Then
                            If Img.GetPixel(1, 55) = Color.FromArgb(255, 1, 0, 0) Then
                                e.Graphics.DrawString("S", myFont2, Branco, 3, 35)
                                e.Graphics.FillEllipse(Laranja, 1, 42, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 51)
                                e.Graphics.DrawString("S", myFont2, Branco, 3, 32)
                                e.Graphics.FillEllipse(Laranja, 1, 39, 10, 10)
                            End If
                        End If
                    ElseIf VGjjj = "FáSustenido" Then
                        If pbox.Name = "P10" OrElse pbox.Name = "P22" OrElse pbox.Name = "P34" OrElse pbox.Name = "P46" OrElse pbox.Name = "P58" OrElse pbox.Name = "P70" OrElse pbox.Name = "P82" Then
                            If Img.GetPixel(1, 55) = Color.FromArgb(255, 1, 0, 0) Then
                                e.Graphics.FillRectangle(TônicaSemitransparente, 0, 0, 17, 85)
                                e.Graphics.DrawString("T", myFont, Vermelho2, 1, 30)
                                e.Graphics.FillEllipse(Vermelho, 1, 42, 10, 10)
                            Else
                                e.Graphics.FillRectangle(TônicaSemitransparente, 0, 0, 17, 85)
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 51)
                                e.Graphics.DrawString("T", myFont, Vermelho2, 1, 27)
                                e.Graphics.FillEllipse(Vermelho, 1, 39, 10, 10)
                            End If
                        ElseIf pbox.Name = "P12" OrElse pbox.Name = "P24" OrElse pbox.Name = "P36" OrElse pbox.Name = "P48" OrElse pbox.Name = "P60" OrElse pbox.Name = "P72" OrElse pbox.Name = "P84" Then
                            If Img.GetPixel(1, 55) = Color.FromArgb(255, 1, 0, 0) Then
                                e.Graphics.DrawString("ST", myFont2, Branco, 1, 35)
                                e.Graphics.FillEllipse(Amarelo, 1, 42, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 51)
                                e.Graphics.DrawString("ST", myFont2, Branco, 1, 32)
                                e.Graphics.FillEllipse(Amarelo, 1, 39, 10, 10)
                            End If
                        ElseIf pbox.Name = "P2" OrElse pbox.Name = "P14" OrElse pbox.Name = "P26" OrElse pbox.Name = "P38" OrElse pbox.Name = "P50" OrElse pbox.Name = "P62" OrElse pbox.Name = "P74" OrElse pbox.Name = "P86" Then
                            If Img.GetPixel(1, 55) = Color.FromArgb(255, 1, 0, 0) Then
                                e.Graphics.DrawString("M", myFont2, Branco, 2, 35)
                                e.Graphics.FillEllipse(Laranja, 1, 42, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 51)
                                e.Graphics.DrawString("M", myFont2, Branco, 2, 32)
                                e.Graphics.FillEllipse(Laranja, 1, 39, 10, 10)
                            End If
                        ElseIf pbox.Name = "P3" OrElse pbox.Name = "P15" OrElse pbox.Name = "P27" OrElse pbox.Name = "P39" OrElse pbox.Name = "P51" OrElse pbox.Name = "P63" OrElse pbox.Name = "P75" OrElse pbox.Name = "P87" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("SD", myFont2, Preto2, 2, 65)
                                e.Graphics.FillEllipse(Rosa, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("SD", myFont2, Preto2, 2, 63)
                                e.Graphics.FillEllipse(Rosa, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P5" OrElse pbox.Name = "P17" OrElse pbox.Name = "P29" OrElse pbox.Name = "P41" OrElse pbox.Name = "P53" OrElse pbox.Name = "P65" OrElse pbox.Name = "P77" OrElse pbox.Name = "T52A" Then
                            If Img.GetPixel(1, 55) = Color.FromArgb(255, 1, 0, 0) Then
                                e.Graphics.DrawString("D", myFont2, Branco, 3, 35)
                                e.Graphics.FillEllipse(Branco, 1, 42, 10, 10)
                                e.Graphics.FillEllipse(Preto, 2, 43, 8, 8)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 51)
                                e.Graphics.DrawString("D", myFont2, Branco, 3, 32)
                                e.Graphics.FillEllipse(Branco, 1, 39, 10, 10)
                                e.Graphics.FillEllipse(Preto, 2, 40, 8, 8)
                            End If
                        ElseIf pbox.Name = "P7" OrElse pbox.Name = "P19" OrElse pbox.Name = "P31" OrElse pbox.Name = "P43" OrElse pbox.Name = "P55" OrElse pbox.Name = "P67" OrElse pbox.Name = "P79" Then
                            If Img.GetPixel(1, 55) = Color.FromArgb(255, 1, 0, 0) Then
                                e.Graphics.DrawString("SpD", myFont2, Branco, -1, 35)
                                e.Graphics.FillEllipse(Verde, 1, 42, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 51)
                                e.Graphics.DrawString("SpD", myFont2, Branco, -1, 32)
                                e.Graphics.FillEllipse(Verde, 1, 39, 10, 10)
                            End If
                        ElseIf pbox.Name = "P8" OrElse pbox.Name = "P20" OrElse pbox.Name = "P32" OrElse pbox.Name = "P44" OrElse pbox.Name = "P56" OrElse pbox.Name = "P68" OrElse pbox.Name = "P80" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("S", myFont2, Preto2, 5, 65)
                                e.Graphics.FillEllipse(Vermelho, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("S", myFont2, Preto2, 5, 63)
                                e.Graphics.FillEllipse(Vermelho, 3, 70, 10, 10)
                            End If
                        End If
                    ElseIf VGjjj = "DóSustenido" Then
                        If pbox.Name = "P5" OrElse pbox.Name = "P17" OrElse pbox.Name = "P29" OrElse pbox.Name = "P41" OrElse pbox.Name = "P53" OrElse pbox.Name = "P65" OrElse pbox.Name = "P77" OrElse pbox.Name = "T52A" Then
                            If Img.GetPixel(1, 55) = Color.FromArgb(255, 1, 0, 0) Then
                                e.Graphics.FillRectangle(TônicaSemitransparente, 0, 0, 17, 85)
                                e.Graphics.DrawString("T", myFont, Vermelho2, 1, 30)
                                e.Graphics.FillEllipse(Branco, 1, 42, 10, 10)
                                e.Graphics.FillEllipse(Preto, 2, 43, 8, 8)
                            Else
                                e.Graphics.FillRectangle(TônicaSemitransparente, 0, 0, 17, 85)
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 51)
                                e.Graphics.DrawString("T", myFont, Vermelho2, 1, 27)
                                e.Graphics.FillEllipse(Branco, 1, 39, 10, 10)
                                e.Graphics.FillEllipse(Preto, 2, 40, 8, 8)
                            End If
                        ElseIf pbox.Name = "P7" OrElse pbox.Name = "P19" OrElse pbox.Name = "P31" OrElse pbox.Name = "P43" OrElse pbox.Name = "P55" OrElse pbox.Name = "P67" OrElse pbox.Name = "P79" Then
                            If Img.GetPixel(1, 55) = Color.FromArgb(255, 1, 0, 0) Then
                                e.Graphics.DrawString("ST", myFont2, Branco, 1, 35)
                                e.Graphics.FillEllipse(Verde, 1, 42, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 51)
                                e.Graphics.DrawString("ST", myFont2, Branco, 1, 32)
                                e.Graphics.FillEllipse(Verde, 1, 39, 10, 10)
                            End If
                        ElseIf pbox.Name = "P8" OrElse pbox.Name = "P20" OrElse pbox.Name = "P32" OrElse pbox.Name = "P44" OrElse pbox.Name = "P56" OrElse pbox.Name = "P68" OrElse pbox.Name = "P80" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("M", myFont2, Preto2, 5, 65)
                                e.Graphics.FillEllipse(Vermelho, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("M", myFont2, Preto2, 5, 63)
                                e.Graphics.FillEllipse(Vermelho, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P10" OrElse pbox.Name = "P22" OrElse pbox.Name = "P34" OrElse pbox.Name = "P46" OrElse pbox.Name = "P58" OrElse pbox.Name = "P70" OrElse pbox.Name = "P82" Then
                            If Img.GetPixel(1, 55) = Color.FromArgb(255, 1, 0, 0) Then
                                e.Graphics.DrawString("SD", myFont2, Branco, 0, 35)
                                e.Graphics.FillEllipse(Vermelho, 1, 42, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 51)
                                e.Graphics.DrawString("SD", myFont2, Branco, 0, 32)
                                e.Graphics.FillEllipse(Vermelho, 1, 39, 10, 10)
                            End If
                        ElseIf pbox.Name = "P12" OrElse pbox.Name = "P24" OrElse pbox.Name = "P36" OrElse pbox.Name = "P48" OrElse pbox.Name = "P60" OrElse pbox.Name = "P72" OrElse pbox.Name = "P84" Then
                            If Img.GetPixel(1, 55) = Color.FromArgb(255, 1, 0, 0) Then
                                e.Graphics.DrawString("D", myFont2, Branco, 3, 35)
                                e.Graphics.FillEllipse(Amarelo, 1, 42, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 51)
                                e.Graphics.DrawString("D", myFont2, Branco, 3, 32)
                                e.Graphics.FillEllipse(Amarelo, 1, 39, 10, 10)
                            End If
                        ElseIf pbox.Name = "P2" OrElse pbox.Name = "P14" OrElse pbox.Name = "P26" OrElse pbox.Name = "P38" OrElse pbox.Name = "P50" OrElse pbox.Name = "P62" OrElse pbox.Name = "P74" OrElse pbox.Name = "P86" Then
                            If Img.GetPixel(1, 55) = Color.FromArgb(255, 1, 0, 0) Then
                                e.Graphics.DrawString("SpD", myFont2, Branco, -1, 35)
                                e.Graphics.FillEllipse(Laranja, 1, 42, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 51)
                                e.Graphics.DrawString("SpD", myFont2, Branco, -1, 32)
                                e.Graphics.FillEllipse(Laranja, 1, 39, 10, 10)
                            End If
                        ElseIf pbox.Name = "P3" OrElse pbox.Name = "P15" OrElse pbox.Name = "P27" OrElse pbox.Name = "P39" OrElse pbox.Name = "P51" OrElse pbox.Name = "P63" OrElse pbox.Name = "P75" OrElse pbox.Name = "P87" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("S", myFont2, Preto2, 5, 65)
                                e.Graphics.FillEllipse(Preto, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("S", myFont2, Preto2, 5, 63)
                                e.Graphics.FillEllipse(Preto, 3, 70, 10, 10)
                            End If
                        End If
                    ElseIf VGjjj = "Fá" Then
                        If pbox.Name = "P9" OrElse pbox.Name = "P21" OrElse pbox.Name = "P33" OrElse pbox.Name = "P45" OrElse pbox.Name = "P57" OrElse pbox.Name = "P69" OrElse pbox.Name = "P81" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.FillRectangle(TônicaSemitransparente, 0, 0, 17, 85)
                                e.Graphics.DrawString("T", myFont, Vermelho2, 3, 60)
                                e.Graphics.FillEllipse(Vermelho, 3, 72, 10, 10)
                            Else
                                e.Graphics.FillRectangle(TônicaSemitransparente, 0, 0, 17, 85)
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("T", myFont, Vermelho2, 3, 58)
                                e.Graphics.FillEllipse(Vermelho, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P11" OrElse pbox.Name = "P23" OrElse pbox.Name = "P35" OrElse pbox.Name = "P47" OrElse pbox.Name = "P59" OrElse pbox.Name = "P71" OrElse pbox.Name = "P83" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("ST", myFont2, Preto2, 3, 65)
                                e.Graphics.FillEllipse(Amarelo, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("ST", myFont2, Preto2, 3, 63)
                                e.Graphics.FillEllipse(Amarelo, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P1" OrElse pbox.Name = "P13" OrElse pbox.Name = "P25" OrElse pbox.Name = "P37" OrElse pbox.Name = "P49" OrElse pbox.Name = "P61" OrElse pbox.Name = "P73" OrElse pbox.Name = "P85" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("M", myFont2, Preto2, 4, 65)
                                e.Graphics.FillEllipse(Laranja, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("M", myFont2, Preto2, 4, 63)
                                e.Graphics.FillEllipse(Laranja, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P2" OrElse pbox.Name = "P14" OrElse pbox.Name = "P26" OrElse pbox.Name = "P38" OrElse pbox.Name = "P50" OrElse pbox.Name = "P62" OrElse pbox.Name = "P74" OrElse pbox.Name = "P86" Then
                            If Img.GetPixel(1, 55) = Color.FromArgb(255, 1, 0, 0) Then
                                e.Graphics.DrawString("SD", myFont2, Branco, 0, 35)
                                e.Graphics.FillEllipse(Rosa, 1, 42, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 51)
                                e.Graphics.DrawString("SD", myFont2, Branco, 0, 32)
                                e.Graphics.FillEllipse(Rosa, 1, 39, 10, 10)
                            End If
                        ElseIf pbox.Name = "P4" OrElse pbox.Name = "P16" OrElse pbox.Name = "P28" OrElse pbox.Name = "P40" OrElse pbox.Name = "P52" OrElse pbox.Name = "P64" OrElse pbox.Name = "P76" OrElse pbox.Name = "P88" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("D", myFont2, Preto2, 5, 65)
                                e.Graphics.FillEllipse(Preto, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("D", myFont2, Preto2, 5, 63)
                                e.Graphics.FillEllipse(Preto, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P6" OrElse pbox.Name = "P18" OrElse pbox.Name = "P30" OrElse pbox.Name = "P42" OrElse pbox.Name = "P54" OrElse pbox.Name = "P66" OrElse pbox.Name = "P78" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("SpD", myFont2, Preto2, 0, 65)
                                e.Graphics.FillEllipse(Verde, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("SpD", myFont2, Preto2, 0, 63)
                                e.Graphics.FillEllipse(Verde, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P8" OrElse pbox.Name = "P20" OrElse pbox.Name = "P32" OrElse pbox.Name = "P44" OrElse pbox.Name = "P56" OrElse pbox.Name = "P68" OrElse pbox.Name = "P80" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("S", myFont2, Preto2, 5, 65)
                                e.Graphics.FillEllipse(Azul, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("S", myFont2, Preto2, 5, 63)
                                e.Graphics.FillEllipse(Azul, 3, 70, 10, 10)
                            End If
                        End If
                    ElseIf VGjjj = "SiBemol" Then
                        If pbox.Name = "P2" OrElse pbox.Name = "P14" OrElse pbox.Name = "P26" OrElse pbox.Name = "P38" OrElse pbox.Name = "P50" OrElse pbox.Name = "P62" OrElse pbox.Name = "P74" OrElse pbox.Name = "P86" Then
                            If Img.GetPixel(1, 55) = Color.FromArgb(255, 1, 0, 0) Then
                                e.Graphics.FillRectangle(TônicaSemitransparente, 0, 0, 17, 85)
                                e.Graphics.DrawString("T", myFont, Vermelho2, 1, 30)
                                e.Graphics.FillEllipse(Rosa, 1, 42, 10, 10)
                            Else
                                e.Graphics.FillRectangle(TônicaSemitransparente, 0, 0, 17, 85)
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 51)
                                e.Graphics.DrawString("T", myFont, Vermelho2, 1, 27)
                                e.Graphics.FillEllipse(Rosa, 1, 39, 10, 10)
                            End If
                        ElseIf pbox.Name = "P4" OrElse pbox.Name = "P16" OrElse pbox.Name = "P28" OrElse pbox.Name = "P40" OrElse pbox.Name = "P52" OrElse pbox.Name = "P64" OrElse pbox.Name = "P76" OrElse pbox.Name = "P88" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("ST", myFont2, Preto2, 3, 65)
                                e.Graphics.FillEllipse(Preto, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("ST", myFont2, Preto2, 3, 63)
                                e.Graphics.FillEllipse(Preto, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P6" OrElse pbox.Name = "P18" OrElse pbox.Name = "P30" OrElse pbox.Name = "P42" OrElse pbox.Name = "P54" OrElse pbox.Name = "P66" OrElse pbox.Name = "P78" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("M", myFont2, Preto2, 4, 65)
                                e.Graphics.FillEllipse(Verde, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("M", myFont2, Preto2, 4, 63)
                                e.Graphics.FillEllipse(Verde, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P7" OrElse pbox.Name = "P19" OrElse pbox.Name = "P31" OrElse pbox.Name = "P43" OrElse pbox.Name = "P55" OrElse pbox.Name = "P67" OrElse pbox.Name = "P79" Then
                            If Img.GetPixel(1, 55) = Color.FromArgb(255, 1, 0, 0) Then
                                e.Graphics.DrawString("SD", myFont2, Branco, 0, 35)
                                e.Graphics.FillEllipse(Azul, 1, 42, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 51)
                                e.Graphics.DrawString("SD", myFont2, Branco, 0, 32)
                                e.Graphics.FillEllipse(Azul, 1, 39, 10, 10)
                            End If
                        ElseIf pbox.Name = "P9" OrElse pbox.Name = "P21" OrElse pbox.Name = "P33" OrElse pbox.Name = "P45" OrElse pbox.Name = "P57" OrElse pbox.Name = "P69" OrElse pbox.Name = "P81" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("D", myFont2, Preto2, 5, 65)
                                e.Graphics.FillEllipse(Vermelho, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("D", myFont2, Preto2, 5, 63)
                                e.Graphics.FillEllipse(Vermelho, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P11" OrElse pbox.Name = "P23" OrElse pbox.Name = "P35" OrElse pbox.Name = "P47" OrElse pbox.Name = "P59" OrElse pbox.Name = "P71" OrElse pbox.Name = "P83" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("SpD", myFont2, Preto2, 0, 65)
                                e.Graphics.FillEllipse(Amarelo, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("SpD", myFont2, Preto2, 0, 63)
                                e.Graphics.FillEllipse(Amarelo, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P1" OrElse pbox.Name = "P13" OrElse pbox.Name = "P25" OrElse pbox.Name = "P37" OrElse pbox.Name = "P49" OrElse pbox.Name = "P61" OrElse pbox.Name = "P73" OrElse pbox.Name = "P85" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("S", myFont2, Preto2, 5, 65)
                                e.Graphics.FillEllipse(Laranja, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("S", myFont2, Preto2, 5, 63)
                                e.Graphics.FillEllipse(Laranja, 3, 70, 10, 10)
                            End If
                        End If
                    ElseIf VGjjj = "MiBemol" Then
                        If pbox.Name = "P2" OrElse pbox.Name = "P14" OrElse pbox.Name = "P26" OrElse pbox.Name = "P38" OrElse pbox.Name = "P50" OrElse pbox.Name = "P62" OrElse pbox.Name = "P74" OrElse pbox.Name = "P86" Then
                            If Img.GetPixel(1, 55) = Color.FromArgb(255, 1, 0, 0) Then
                                e.Graphics.FillRectangle(TônicaSemitransparente, 0, 0, 17, 85)
                                e.Graphics.DrawString("T", myFont, Vermelho2, 1, 30)
                                e.Graphics.FillEllipse(Rosa, 1, 42, 10, 10)
                            Else
                                e.Graphics.FillRectangle(TônicaSemitransparente, 0, 0, 17, 85)
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 51)
                                e.Graphics.DrawString("T", myFont, Vermelho2, 1, 27)
                                e.Graphics.FillEllipse(Rosa, 1, 39, 10, 10)
                            End If
                        ElseIf pbox.Name = "P4" OrElse pbox.Name = "P16" OrElse pbox.Name = "P28" OrElse pbox.Name = "P40" OrElse pbox.Name = "P52" OrElse pbox.Name = "P64" OrElse pbox.Name = "P76" OrElse pbox.Name = "P88" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("ST", myFont2, Preto2, 3, 65)
                                e.Graphics.FillEllipse(Preto, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("ST", myFont2, Preto2, 3, 63)
                                e.Graphics.FillEllipse(Preto, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P6" OrElse pbox.Name = "P18" OrElse pbox.Name = "P30" OrElse pbox.Name = "P42" OrElse pbox.Name = "P54" OrElse pbox.Name = "P66" OrElse pbox.Name = "P78" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("M", myFont2, Preto2, 4, 65)
                                e.Graphics.FillEllipse(Verde, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("M", myFont2, Preto2, 4, 63)
                                e.Graphics.FillEllipse(Verde, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P7" OrElse pbox.Name = "P19" OrElse pbox.Name = "P31" OrElse pbox.Name = "P43" OrElse pbox.Name = "P55" OrElse pbox.Name = "P67" OrElse pbox.Name = "P79" Then
                            If Img.GetPixel(1, 55) = Color.FromArgb(255, 1, 0, 0) Then
                                e.Graphics.DrawString("SD", myFont2, Branco, 0, 35)
                                e.Graphics.FillEllipse(Azul, 1, 42, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 51)
                                e.Graphics.DrawString("SD", myFont2, Branco, 0, 32)
                                e.Graphics.FillEllipse(Azul, 1, 39, 10, 10)
                            End If
                        ElseIf pbox.Name = "P9" OrElse pbox.Name = "P21" OrElse pbox.Name = "P33" OrElse pbox.Name = "P45" OrElse pbox.Name = "P57" OrElse pbox.Name = "P69" OrElse pbox.Name = "P81" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("D", myFont2, Preto2, 5, 65)
                                e.Graphics.FillEllipse(Vermelho, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("D", myFont2, Preto2, 5, 63)
                                e.Graphics.FillEllipse(Vermelho, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P11" OrElse pbox.Name = "P23" OrElse pbox.Name = "P35" OrElse pbox.Name = "P47" OrElse pbox.Name = "P59" OrElse pbox.Name = "P71" OrElse pbox.Name = "P83" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("SpD", myFont2, Preto2, 0, 65)
                                e.Graphics.FillEllipse(Amarelo, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("SpD", myFont2, Preto2, 0, 63)
                                e.Graphics.FillEllipse(Amarelo, 3, 70, 10, 10)
                            End If
                        ElseIf pbox.Name = "P1" OrElse pbox.Name = "P13" OrElse pbox.Name = "P25" OrElse pbox.Name = "P37" OrElse pbox.Name = "P49" OrElse pbox.Name = "P61" OrElse pbox.Name = "P73" OrElse pbox.Name = "P85" Then
                            If Img.GetPixel(1, 84) = Color.FromArgb(255, 194, 211, 220) Then
                                e.Graphics.DrawString("S", myFont2, Preto2, 5, 65)
                                e.Graphics.FillEllipse(Laranja, 3, 72, 10, 10)
                            Else
                                If VGhhh = 1 Then e.Graphics.FillRectangle(FundoSemitransparente, 0, 0, 17, 82)
                                e.Graphics.DrawString("S", myFont2, Preto2, 5, 63)
                                e.Graphics.FillEllipse(Laranja, 3, 70, 10, 10)
                            End If
                        End If
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

    Private Sub IdentificaNomeMenu(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem17.Click, ToolStripMenuItem16.Click, ChinTTTTTToolStripMenuItem.Click, ChiaoToolStripMenuItem.Click, ChaioTTTTTToolStripMenuItem.Click, Bizantina2ToolStripMenuItem.Click, Bizantina1ToolStripMenuItem.Click, _
    BiYuT2TTTToolStripMenuItem.Click, BeBopSemidiminuídasTTsssTsToolStripMenuItem.Click, BeBopMenorTsssTTsTToolStripMenuItem.Click, BeBopMaiorTTsTssTsToolStripMenuItem.Click, BeBopDominanteTTsTTsssToolStripMenuItem.Click, Balinesa2ToolStripMenuItem.Click, Balinesa1sTTT2TToolStripMenuItem.Click, AuxiliarAumentadaToolStripMenuItem.Click, Aumentada3ToolStripMenuItem.Click, Aumentada2TsTsTTToolStripMenuItem.Click, Aumentada1TsTsTsToolStripMenuItem.Click, Arábe5sTssTsTssToolStripMenuItem.Click, Árabe4TsTTTsssToolStripMenuItem.Click, Árabe3TsTsTsTToolStripMenuItem.Click, Árabe2TsTsTsTsToolStripMenuItem.Click, Árabe1TTssTTTToolStripMenuItem.Click, AlhijazToolStripMenuItem.Click, _
    AhavohRabbohToolStripMenuItem.Click, ToolStripMenuItem11.Click, TonsInteirosDiminuídaDiminishedToolStripMenuItem.Click, HispanoárabeToolStripMenuItem.Click, Etíope3TTsTsTsToolStripMenuItem.Click, Etíope2OuEthiopianGeezEzelToolStripMenuItem.Click, Etíope1ToolStripMenuItem.Click, EsquimalTetratônicaTTT2TToolStripMenuItem.Click, EsquimalHexatónica2ToolStripMenuItem.Click, EsquimalHexatônica1TTTTTsToolStripMenuItem.Click, EsquimalHeptatônicaToolStripMenuItem.Click, EsplasTsssTTTToolStripMenuItem.Click, EspanholaOctatônicasTsssTTTToolStripMenuItem.Click, EspanholaHexatônicasTsTTTToolStripMenuItem.Click, EnigmáticaDeVerdi3sTTTTssToolStripMenuItem.Click, _
    EnigmáticaDeVerdi2sTsTTssToolStripMenuItem.Click, EnigmáticaDeVerdi1sTssTTssToolStripMenuItem.Click, Enigmática2sTTTTsToolStripMenuItem.Click, Enigmática1sTTT2TToolStripMenuItem.Click, EgípciaToolStripMenuItem.Click, DiminuídaHalfToolStripMenuItem.Click, Diminuída3ToolStripMenuItem.Click, Coreana2TTTTsTToolStripMenuItem.Click, Coreana1ToolStripMenuItem.Click, CíngaraMenorTsTssTsToolStripMenuItem.Click, CíngaraMaior2sTTssTTToolStripMenuItem.Click, CíngaraMaior1ToolStripMenuItem.Click, CíngaraHexatônicasTsTssTToolStripMenuItem.Click, CingaraEspanholaToolStripMenuItem.Click, CiganaHungaroPersaMenuItem.Click, CiganaEspanholaToolStripMenuItem1.Click, _
    ChinaOctatônicaTTsTTsssToolStripMenuItem.Click, ChinaAntiguaTTTsTTToolStripMenuItem.Click, China2ToolStripMenuItem.Click, China12TTs2TsToolStripMenuItem.Click, AuxiliarDiminuídaBluessTsTsTsTToolStripMenuItem.Click, FlamencaToolStripMenuItem.Click, JinYuToolStripMenuItem.Click, JavanesaToolStripMenuItem.Click, Javanesa3ToolStripMenuItem.Click, Iwatos2Ts2TTToolStripMenuItem.Click, Israelita2ToolStripMenuItem.Click, IshikotsuchoOuIchikosuchoTTsssTTsToolStripMenuItem.Click, InToolStripMenuItem.Click, HúngaraMenor2TsTssTsToolStripMenuItem.Click, HúngaraMenor1TsTssTTToolStripMenuItem.Click, HúngaraMaior2sTTssTTToolStripMenuItem.Click, HúngaraMaior1TsTsTsTToolStripMenuItem.Click, _
    HouzamTssTTTsToolStripMenuItem.Click, Honkumoijoshis2TTs2TToolStripMenuItem.Click, HonchoshiPlagalsTTs2TTToolStripMenuItem.Click, Honchoshi2T3TToolStripMenuItem.Click, HitzazMenuItem.Click, HitzaskiarMenuItem.Click, HirajoshiTs2Ts2TToolStripMenuItem.Click, HedjazTsTsTsTToolStripMenuItem.Click, Hawayana2ToolStripMenuItem.Click, Hawayana1Ts2TTTsToolStripMenuItem.Click, HanKumoiTTTs2TToolStripMenuItem.Click, GregorianaTsTTsssTToolStripMenuItem.Click, GongToolStripMenuItem.Click, GhanaPentatônica2ToolStripMenuItem.Click, GhanaPentatônica1ToolStripMenuItem.Click, GhanaHeptatônicaToolStripMenuItem.Click, GenusTertiumTsTsTsToolStripMenuItem.Click, _
    GenusSecundum2TsTTTsToolStripMenuItem.Click, GenusPrimumTTT2TToolStripMenuItem.Click, GenusDiatonicumVeterumTTsssTTsToolStripMenuItem.Click, GenusDiatonicumTTsTTsssToolStripMenuItem.Click, GenusChromaticumsTssTssTsToolStripMenuItem.Click, ToolStripMenuItem12.Click, OusakToolStripMenuItem.Click, Oriental2sTssTsTToolStripMenuItem.Click, Oriental1sTssTTTToolStripMenuItem.Click, NohkanTTsTsTsToolStripMenuItem.Click, NiaventTsTssTsToolStripMenuItem.Click, NiagariHexatônicas2TTsTTToolStripMenuItem.Click, NiagariDitônica3T2TToolStripMenuItem.Click, NapolitanaMenorToolStripMenuItem.Click, Napolitana3sTTTsTsToolStripMenuItem.Click, Napolitana2sTTTTTsToolStripMenuItem.Click, _
    Napolitana1sTTTsTToolStripMenuItem.Click, MongolChinesaToolStripMenuItem.Click, Mischung6TTsTsTTToolStripMenuItem.Click, Mischung5ToolStripMenuItem.Click, Mischung4ToolStripMenuItem.Click, Mischung3ToolStripMenuItem.Click, Mischung2TTsTsTsToolStripMenuItem.Click, Mischung1ToolStripMenuItem.Click, Messiânica5ssssTssssTToolStripMenuItem.Click, Messiânica4ssTTssTTToolStripMenuItem.Click, Messiânica3ss2Tss2TToolStripMenuItem.Click, Messiânica2sssTsssTToolStripMenuItem.Click, Messiânica1ssTsssTsTToolStripMenuItem.Click, MelaYagapriya31TssTssTToolStripMenuItem.Click, MelaVisvambari54TsTsTssToolStripMenuItem.Click, MelaVaschaspati64TTTsTsTToolStripMenuItem.Click, _
    MelaVarunapriya24TsTTTssToolStripMenuItem.Click, MelaVanaspati4ssTTTsTToolStripMenuItem.Click, MelaVakulabharanamToolStripMenuItem.Click, MelaVagadhisvari34TssTTsTToolStripMenuItem.Click, MelaTanarupi6ssTTTssToolStripMenuItem.Click, MelaSyamalangi55TsTsssTToolStripMenuItem.Click, MelaSuvarnangi47ss2TsTTsToolStripMenuItem.Click, MelaSuryakantam17sTsTTTsToolStripMenuItem.Click, MelaSulini35TssTTTsToolStripMenuItem.Click, MelaSucharitra67TsTsssTToolStripMenuItem.Click, MelaSubhapantuvarali45sTTssTsToolStripMenuItem.Click, MelaSimhendramadhyama57TsTssTsToolStripMenuItem.Click, MelaSenavati7sTTTssTToolStripMenuItem.Click, MelaSarasangi27TTsTsTsToolStripMenuItem.Click, _
    MelaSanmukhapriya56TsTssTTToolStripMenuItem.Click, MelaSalagam37Scaless2TsssTToolStripMenuItem.Click, MelaSadvidhamargini46sTTsTsTToolStripMenuItem.Click, MelaRupavati12sTTTTssToolStripMenuItem.Click, MelaRisabhapriya62TTTssTTToolStripMenuItem.Click, MelaRatnangi2ssTTsTTToolStripMenuItem.Click, MelaRasikapriya72TsTsTssToolStripMenuItem.Click, MelaRamapriya52sTTsTsTToolStripMenuItem.Click, MelaRaghupriya42ss2TsTssToolStripMenuItem.Click, MelaRagavardhani32ScaleToolStripMenuItem.Click, MelaPavani41ss2TsTTsToolStripMenuItem.Click, MelaNitimati60TsTsTssToolStripMenuItem.Click, MelaNavanitam40ss2TsTsTToolStripMenuItem.Click, MelaNatakapriya10ToolStripMenuItem.Click, _
    MelaNatabhairavi20ToolStripMenuItem.Click, MelaNasikabhusani70TsTsTsTToolStripMenuItem.Click, MelaNamanarayani50sTTssTTToolStripMenuItem.Click, MelaNaganandini30TTsTTssToolStripMenuItem.Click, MelaMechakalyani65ToolStripMenuItem.Click, MelaMayamalavagaula15sTsTssTToolStripMenuItem.Click, MelaMararanjani25TTsTssTToolStripMenuItem.Click, MelaManavati5ssTTTTsToolStripMenuItem.Click, MelaLatangi63TTTssTsToolStripMenuItem.Click, MelaKosalam71TsTsTTsToolStripMenuItem.Click, MelaKokilapriya11sTTTTTsToolStripMenuItem.Click, MelaKiravani21ToolStripMenuItem.Click, MelaKharaharapriya22ToolStripMenuItem.Click, MelaKantamani61TTTsssTToolStripMenuItem.Click, _
    MelaKanakangi1ToolStripMenuItem.Click, MelaKamavardhani51sTTssTsToolStripMenuItem.Click, MelaJyotisvarupini68TsTssTTToolStripMenuItem.Click, MelaJhankaradhvani19TsTTssTToolStripMenuItem.Click, MelaJhalavarali39ss2TssTsToolStripMenuItem.Click, MelaJalarnavam38ss2TssTTToolStripMenuItem.Click, MelaHemavati58TsTsTsTToolStripMenuItem.Click, MelaHatakambari18sTsTTssToolStripMenuItem.Click, MelaHarikambhoji28ToolStripMenuItem.Click, MelaHanumattodi8ToolStripMenuItem.Click, MelaGayakapriya13sTsTssTToolStripMenuItem.Click, MelaGavambodhi43sTTsssTToolStripMenuItem.Click, MelaGaurimanohari23ToolStripMenuItem.Click, MelaGangeyabhusani33TssTsTsToolStripMenuItem.Click, _
    MelaGanamurti3ssTTsTsToolStripMenuItem.Click, MelaGamanasrama53sTTsTTsToolStripMenuItem.Click, MelaDivyamani48sTTsTssToolStripMenuItem.Click, MelaDhirasankarabharana29ToolStripMenuItem.Click, MelaDhenuka9sTTTsTsToolStripMenuItem.Click, MelaDhavalambari49sTTsssTToolStripMenuItem.Click, MelaDhatuvardhani69TsTssTsToolStripMenuItem.Click, MelaDharmavati59TsTsTTsToolStripMenuItem.Click, MelaChitrambari66TTTsTssToolStripMenuItem.Click, _
MelaCharukesi26TTsTsTTToolStripMenuItem.Click, MelaChalanata36TssTTssToolStripMenuItem.Click, MelaChakravakam16sTsTTsTToolStripMenuItem.Click, MelaBhavapriya44ToolStripMenuItem.Click, MaqamZenguleToolStripMenuItem.Click, MaqamSuzdilTsTssTTToolStripMenuItem.Click, MaqamShawqTTsTTsssToolStripMenuItem.Click, MaqamShahnazKurdisTTTsTsToolStripMenuItem.Click, MaqamShaddarabansTsssTsTToolStripMenuItem.Click, MaqamNakrizTsTsTsTToolStripMenuItem.Click, MaqamNahawandTsTTsTssToolStripMenuItem.Click, MaqamKurdToolStripMenuItem.Click, MaqamKarcigarTsTsTsTToolStripMenuItem.Click, MaqamHuzzamsTsTsTTToolStripMenuItem.Click, MaqamHumayunToolStripMenuItem.Click, _
MaqamHijazsTsTsTssToolStripMenuItem.Click, MaqamHicazsTsTTsTToolStripMenuItem.Click, MaqamBayateEsfahanToolStripMenuItem.Click, MaometanaToolStripMenuItem.Click, MagenAbotsTsTTsTsToolStripMenuItem.Click, KyemyonjoTTTTTToolStripMenuItem.Click, KungTTTTTToolStripMenuItem.Click, KumoiScaleToolStripMenuItem.Click, Kokinjoshis2TTTTToolStripMenuItem.Click, JudaicaAhabaRabbaToolStripMenuItem.Click, RagaLavangis3TTTToolStripMenuItem.Click, RagaLatikaTTTsTsToolStripMenuItem.Click, RagaLalitasTssTTsToolStripMenuItem.Click, RagaKuntvarali2TTTsTToolStripMenuItem.Click, RagaKumurdakiTTT2TsToolStripMenuItem.Click, RagaKumudTTTTTsToolStripMenuItem.Click, _
RagaKshanikas2TTTsToolStripMenuItem.Click, RagaKokilPanchamTTTs2TToolStripMenuItem.Click, RagaKiranavaliToolStripMenuItem.Click, RagaKhamas2TsTTsTToolStripMenuItem.Click, RagaKhamajiDurga2Ts2TsTToolStripMenuItem.Click, RagaKanakambariToolStripMenuItem.Click, RagaKambhojiTTsTTTToolStripMenuItem.Click, RagaKamalamanohari2TsTsTTToolStripMenuItem.Click, RagaKalavatisTsTTTToolStripMenuItem.Click, RagaKalakanthis2TTssTToolStripMenuItem.Click, RagaKalagadasTTssTToolStripMenuItem.Click, RagaKaikavasiToolStripMenuItem.Click, RagaJyoti2TTssTTToolStripMenuItem.Click, RagaJivantiniTTsTssToolStripMenuItem.Click, RagaJivantikas2TTTTsToolStripMenuItem.Click, RagaJayakaunsTTs2TTToolStripMenuItem.Click, RagaJaganmohanamT2TssTTToolStripMenuItem.Click, RagaHindol2TTTTsToolStripMenuItem.Click, _
RagaHejjajjisTTTsTToolStripMenuItem.Click, RagaHariNata2TsTTTsToolStripMenuItem.Click, RagaHarikaunsTTTTTToolStripMenuItem.Click, RagaHamsaVinodiniTTs2TTsToolStripMenuItem.Click, RagaHamsanandisTTTTsToolStripMenuItem.Click, RagaHamsadhvaniTTTTTToolStripMenuItem.Click, RagaGurjariTodisTTTTsToolStripMenuItem.Click, RagaGorakhKalyanTT2TsTToolStripMenuItem.Click, RagaGopriyaToolStripMenuItem.Click, RagaGopikavasantamTTTsTTToolStripMenuItem.Click, RagaGirija2TsTTsToolStripMenuItem.Click, RagaGhantanaToolStripMenuItem.Click, RagaGauris2TT2TsToolStripMenuItem.Click, RagaGaulasTsT2TsToolStripMenuItem.Click, RagaGandharavamsTTTTTToolStripMenuItem.Click, RagaGanasamavaralissTTsTsToolStripMenuItem.Click, RagaGambhiranata2TsT2TsToolStripMenuItem.Click, RagaDipakTTsss2TToolStripMenuItem.Click, _
RagaDhavalashri2TTsTTToolStripMenuItem.Click, RagaDhavalangamsTTss2TToolStripMenuItem.Click, RagaDevranjani2TTsTTToolStripMenuItem.Click, RagaDevaranji2TTsTsToolStripMenuItem.Click, RagaDevakriyaTTTTTToolStripMenuItem.Click, RagaDeshTTT2TsToolStripMenuItem.Click, RagaDeshgaurs3TsTsToolStripMenuItem.Click, RagaDarbarTTTTsTToolStripMenuItem.Click, RagaCintamaniToolStripMenuItem.Click, RagaChhayaTodisTTT2TToolStripMenuItem.Click, RagaChandrakaunsModernTT2TTsToolStripMenuItem.Click, RagaChandrakaunsKiravaniTTTTsToolStripMenuItem.Click, RagaChandrakaunsKafiTT2TsTToolStripMenuItem.Click, RagaChandrajyotiss2TsTTToolStripMenuItem.Click, RagaCaturanginiTTTs2TsToolStripMenuItem.Click, _
RagaBrindabaniSarangTTTTssToolStripMenuItem.Click, RagaBilashkhaniTodiToolStripMenuItem.Click, RagaBhupeshwariTTTs2TToolStripMenuItem.Click, RagaBhupalamToolStripMenuItem.Click, RagaBhinnaShadja2Ts2TTsToolStripMenuItem.Click, RagaBhinnaPancamaTTTsTsToolStripMenuItem.Click, RagaBhavaniTetratônicaTT2TTToolStripMenuItem.Click, RagaBhavaniHexatônicasTTTTTToolStripMenuItem.Click, RagaBhatiyarsTsssTTsToolStripMenuItem.Click, RagaBhanumatissTTTsTToolStripMenuItem.Click, RagaBhanumanjariTssTTTToolStripMenuItem.Click, RagaBaulisTTsTsToolStripMenuItem.Click, RagaBarbaraTTTTsTToolStripMenuItem.Click, RagaBagesriToolStripMenuItem.Click, RagaAudavTukhariToolStripMenuItem.Click, RagaAmritavarsini2TTs2TsToolStripMenuItem.Click, RagaAmarasenapriyaToolStripMenuItem.Click, RagaAhirBhairavsTsTTsTToolStripMenuItem.Click, _
RagaAdanaToolStripMenuItem.Click, RagaAbhogiToolStripMenuItem.Click, PyongjoTTTTsTToolStripMenuItem.Click, PienChihToolStripMenuItem.Click, PeruanaTritônica2T2T2TToolStripMenuItem.Click, PeruanaTritônica12TT2TToolStripMenuItem.Click, PeruanaMenorToolStripMenuItem.Click, PeruanaMaiorToolStripMenuItem.Click, Persa2ToolStripMenuItem.Click, Persa1sTssTTsToolStripMenuItem.Click, PelogToolStripMenuItem.Click, PeiraiotikossTTsTTsToolStripMenuItem.Click, YuPentatônicaToolStripMenuItem.Click, YuHeptatônicaToolStripMenuItem.Click, YoulanssTsssTsTToolStripMenuItem.Click, YoToolStripMenuItem.Click, YosenTTTTsTToolStripMenuItem.Click, YiZeTTTTTToolStripMenuItem.Click, _
WaraoTritônica2TT2TToolStripMenuItem.Click, WaraoTetratônicaTs3TTToolStripMenuItem.Click, WaraoDitônicasTToolStripMenuItem.Click, UteT3TTToolStripMenuItem.Click, UjoTTTTTToolStripMenuItem.Click, TonalToolStripMenuItem.Click, TodiThetasTTssTsToolStripMenuItem.Click, ThetaKhamajToolStripMenuItem.Click, ThetaKalyanToolStripMenuItem.Click, ThetaKafiScaleToolStripMenuItem.Click, ThetaBilavalToolStripMenuItem.Click, ThetaBhairavToolStripMenuItem.Click, ThetaBhairaviToolStripMenuItem.Click, ThetaAsavariToolStripMenuItem.Click, TcherepninsTssTssTsToolStripMenuItem.Click, TaishikichoTTsssTsssToolStripMenuItem.Click, SouzinakTsTsTsTToolStripMenuItem.Click, _
SiriasTsT2TToolStripMenuItem.Click, SimétricaHexatônicasTsTsTToolStripMenuItem.Click, Simétrica3ssTssssTssToolStripMenuItem.Click, Simétrica2ToolStripMenuItem.Click, ShangTTTTTToolStripMenuItem.Click, SengahTssTsTsToolStripMenuItem.Click, Sansagari2T2TTToolStripMenuItem.Click, SambahTssTsTTToolStripMenuItem.Click, Ryukyu2TsT2TsToolStripMenuItem.Click, RyosenToolStripMenuItem.Click, RomanaMenorTsTsTsTToolStripMenuItem.Click, RitusenTTTTTToolStripMenuItem.Click, RitsusTTTTTToolStripMenuItem.Click, RastTTsTTsssToolStripMenuItem.Click, RagaZilaf2TsTs2TToolStripMenuItem.Click, _
RagaYamunaKalyaniTTTsTTToolStripMenuItem.Click, RagaVutari2TTsTsTToolStripMenuItem.Click, RagaViyogavaralisTTTTsToolStripMenuItem.Click, RagaVijayavasanta2TTsTssToolStripMenuItem.Click, RagaVijayasrisTTs2TsToolStripMenuItem.Click, RagaVijayanagariToolStripMenuItem.Click, RagaVibhavaris2TTTTToolStripMenuItem.Click, RagaVasantasTs2TTsToolStripMenuItem.Click, RagaVasantabhairavisTsTTTToolStripMenuItem.Click, RagaValaji2TTTsTToolStripMenuItem.Click, RagaVaijayantiT2Ts2TsToolStripMenuItem.Click, RagaTrimurtiToolStripMenuItem.Click, RagaTilang2TsTTssToolStripMenuItem.Click, RagaTakkaTTTsTsToolStripMenuItem.Click, RagaSyamalamToolStripMenuItem.Click, _
RagaSumukamT2T2TsToolStripMenuItem.Click, RagaSuddhaTodisTTTTTToolStripMenuItem.Click, RagaSuddhaSimantinisTTTs2TToolStripMenuItem.Click, RagaSuddhaMukharissTTsTToolStripMenuItem.Click, RagaSuddhaBangalaToolStripMenuItem.Click, RagaSoratiTTTTsssToolStripMenuItem.Click, RagaSivaranjiniToolStripMenuItem.Click, RagaSivaKambhojiTTsTTTToolStripMenuItem.Click, RagaSindhuraKafiToolStripMenuItem.Click, RagaSindhiBhairaviToolStripMenuItem.Click, RagaSimharavaToolStripMenuItem.Click, RagaShuddhKalyanToolStripMenuItem.Click, RagaShrisTTssTsToolStripMenuItem.Click, RagaShriKalyanT2TsTTToolStripMenuItem.Click, RagaShobhavariTTTs2TToolStripMenuItem.Click, RagaShailajaT2TsTTToolStripMenuItem.Click, RagaSaurastrasTsTssTsToolStripMenuItem.Click, RagaSaugandhinis2Tss2TToolStripMenuItem.Click, _
RagaSarvasri2TT2TToolStripMenuItem.Click, RagaSaravati2TsTssTToolStripMenuItem.Click, RagaSarasvatiT2TsTsTToolStripMenuItem.Click, RagaSarasananaTTsTTsToolStripMenuItem.Click, RagaSarangTTTTssToolStripMenuItem.Click, RagaSamudhraPriyaTTsTTToolStripMenuItem.Click, RagaSalanganatas2TTs2TToolStripMenuItem.Click, RagaSalagavaralisT2TTsTToolStripMenuItem.Click, RagaRudraPancamasTs2TsTToolStripMenuItem.Click, RagaRevasTTs2TToolStripMenuItem.Click, RagaRasranjaniTT2TTsToolStripMenuItem.Click, RagaRasikaRanjanisTTTTToolStripMenuItem.Click, RagaRasavalis2TTTsTToolStripMenuItem.Click, RagaRasamanjariTsTs2TsToolStripMenuItem.Click, RagaRanjaniToolStripMenuItem.Click, RagaRamkaliToolStripMenuItem.Click, RagaRamdasiMalharTssTTssssToolStripMenuItem.Click, RagaRagesriTTs2TsssToolStripMenuItem.Click, RagaRageshriTTs2TsTToolStripMenuItem.Click, RagaPuruhutika2TTTTsToolStripMenuItem.Click, RagaPurnaPancamasTsTs2TToolStripMenuItem.Click, _
RagaPurnalalitaToolStripMenuItem.Click, RagaPriyadharshiniTTTTsToolStripMenuItem.Click, RagaPiluTsTTsssssToolStripMenuItem.Click, RagaPhenadyutis2TTsTTToolStripMenuItem.Click, RagaPatdipToolStripMenuItem.Click, RagaParaju2TsTsTsToolStripMenuItem.Click, RagaPalasiToolStripMenuItem.Click, RagaPadis2TTsTsToolStripMenuItem.Click, RagaOngkari3Ts2TToolStripMenuItem.Click, RagaNeroshtaTT2TTTToolStripMenuItem.Click, RagaNavamanohariTTTsTTToolStripMenuItem.Click, RagaNataTTT2TsToolStripMenuItem.Click, RagaNalinakantiTTsT2TsToolStripMenuItem.Click, RagaNagasvaravali2TsTTTToolStripMenuItem.Click, RagaNagagandhariTTTTTsToolStripMenuItem.Click, RagaNabhomaniss2Ts2TToolStripMenuItem.Click, RagaMultaniTTs2TsToolStripMenuItem.Click, RagaMukhariTsTTsssTToolStripMenuItem.Click, _
RagaMruganandanaTTTTTsToolStripMenuItem.Click, RagaMohanangiTsTTTToolStripMenuItem.Click, RagaMianKiMalharTsTTTsssToolStripMenuItem.Click, RagaMegharanjisTs3TsToolStripMenuItem.Click, RagaMegharanjanisTsT2TToolStripMenuItem.Click, RagaMathaKokilaT2TTsTToolStripMenuItem.Click, RagaManohariTTTTsTToolStripMenuItem.Click, RagaMandarisTTs2TsToolStripMenuItem.Click, RagaManaviToolStripMenuItem.Click, RagaManaranjaniIsTTTTToolStripMenuItem.Click, RagaManaranjaniIIs2TTTTToolStripMenuItem.Click, RagaMamata2TTTTsToolStripMenuItem.Click, RagaMalkaunsTTTTTToolStripMenuItem.Click, RagaMalayamarutamsTTTsTToolStripMenuItem.Click, RagaMalasri2TT2TToolStripMenuItem.Click, RagaMahathi2TTTTToolStripMenuItem.Click, RagaMadhyamavatiToolStripMenuItem.Click, RagaMadhuri2TsTTsssToolStripMenuItem.Click, RagaMadhukaunsTTsTsTToolStripMenuItem.Click, ZhiTTTTTToolStripMenuItem.Click, UltraLócriasTsTTsTToolStripMenuItem.Click, ToolStripMenuItem13.Click, _
ToolStripMenuItem1.Click, TesteToolStripMenuItem5.Click, TesteToolStripMenuItem3.Click, SuperLócriosTsTTTTToolStripMenuItem.Click, PentatônicaNeutraToolStripMenuItem.Click, PentatônicaNeutra2s2TTTTToolStripMenuItem.Click, PentatônicaMenor4Ts2TTTToolStripMenuItem.Click, PentatônicaMenor3TTs2TTToolStripMenuItem.Click, PentatônicaMenor2ToolStripMenuItem.Click, PentatônicaMaior2TTT2TsToolStripMenuItem.Click, PentatônicaDeDominanteTTTTTToolStripMenuItem.Click, PentatônicaBluesTTssTTToolStripMenuItem.Click, PentatônicaBluesToolStripMenuItem.Click, PentatônicaAltB6TTTs2TToolStripMenuItem.Click, PentatônicaAltB5TTTTTToolStripMenuItem.Click, PentatônicaAltb3b6Ts2Ts2TToolStripMenuItem.Click, PentatônicaAltb2sTTTTToolStripMenuItem.Click, OctatônicasHWToolStripMenuItem.Click, MixolídioTTSTTSTToolStripMenuItem.Click, MixolidiaHexatônicaTTTTsTToolStripMenuItem.Click, MixolidiaCromáticassTssTTToolStripMenuItem.Click, MixolidiaAumentadaTTsTssTToolStripMenuItem.Click, MenorToolStripMenuItem.Click, MenorMelódicaToolStripMenuItem.Click, MenorMelódicaDescendenteTSTTSTTMenorNaturalToolStripMenuItem.Click, MenorHexatônicaToolStripMenuItem.Click, MenorHarmônicaTsTTsTsToolStripMenuItem.Click, MenorHarmônicaTSTTST12ToolStripMenuItem.Click, _
MaiorToolStripMenuItem.Click, LócrioSTTSTTTToolStripMenuItem.Click, LocriaMayorTTssTTTToolStripMenuItem.Click, LídioTTTSTTSToolStripMenuItem.Click, LidiaMenorScaleTTTssTTToolStripMenuItem.Click, LidiaHexatônicaTTTTTsToolStripMenuItem.Click, LidiaDiminuídaTsTsTTsToolStripMenuItem.Click, LidiaCromáticasTssTTsToolStripMenuItem.Click, LidiaB7TTTsTsTToolStripMenuItem.Click, LidiaAumentadaTTTTsTsToolStripMenuItem.Click, JônioTTSTTTSToolStripMenuItem.Click, JônioAumentadaTTsTsTsToolStripMenuItem.Click, JazzMenorToolStripMenuItem.Click, HipolidiaCromáticasTTssTsToolStripMenuItem.Click, HipofrigiaCromáticaToolStripMenuItem.Click, HipodóricaCromáticaTssTssTToolStripMenuItem.Click, FrígioSTTTSTTToolStripMenuItem.Click, FrigioMaiorToolStripMenuItem.Click, FrigiaHexatônicaTTTsTTToolStripMenuItem.Click, FrigiaEspanholasTssTsTTToolStripMenuItem.Click, FrigiaDobleHexatônicasTTsTTToolStripMenuItem.Click, FrigiaCromáticaTssTssTToolStripMenuItem.Click, FrigiaB4sTsTsTTToolStripMenuItem.Click, FrigiaÁrabesTssTsTssToolStripMenuItem.Click, Frigia6ToolStripMenuItem.Click, _
EólioTSTTSTTModoMenorToolStripMenuItem.Click, DuplaHarmônicaToolStripMenuItem.Click, DórioAlteradaTsTsTTTToolStripMenuItem.Click, DóricoTSTTTSTToolStripMenuItem.Click, DóricaCromáticaToolStripMenuItem.Click, DóricaB2ToolStripMenuItem.Click, DiatônicaToolStripMenuItem.Click, BluesMenorTTssTTToolStripMenuItem.Click, BluesMaiorToolStripMenuItem.Click, Blues9ToolStripMenuItem.Click, Blues8ToolStripMenuItem.Click, Blues7ToolStripMenuItem.Click, Blues6ToolStripMenuItem.Click, Blues5ToolStripMenuItem.Click, Blues4ToolStripMenuItem.Click, Blues3ToolStripMenuItem.Click, Blues2ToolStripMenuItem.Click, Blues1ToolStripMenuItem.Click, Blues10ToolStripMenuItem.Click, AuxiliarDiminuídaBluesToolStripMenuItem.Click, TonsInteirosTTTTTTToolStripMenuItem.Click, TonsInteirosDiminuídasTsTTTTToolStripMenuItem.Click, SemitomTomToolStripMenuItem.Click, _
LeadingWholeToneToolStripMenuItem.Click, TritôToolStripMenuItem.Click, Tetratônica2Ts2TsToolStripMenuItem.Click, Diminuída2TsTsTsTToolStripMenuItem.Click, DomSus4TTTTsTToolStripMenuItem.Click, JaponesaBScaleTTTs2TToolStripMenuItem.Click, Skriabin1TTTTsTToolStripMenuItem1.Click, PrometheusTTTTsTToolStripMenuItem.Click, JudaicaTTsTsTsToolStripMenuItem.Click, HarmônicaMaiorTTsTsTsToolStripMenuItem.Click, HinduHindustanOuIndustanToolStripMenuItem.Click, HexatônicaPiramidalTsTsTTToolStripMenuItem.Click, SemidiminuídaTsTsTTTToolStripMenuItem.Click, Diminuída1TsTsTsTsToolStripMenuItem.Click, Simétrica1TsTsTsTsToolStripMenuItem.Click, TomsemitomTsTsTsTsToolStripMenuItem.Click, Semidiminuída2Lócrio2TsTsTTTToolStripMenuItem1.Click, Semidiminuída2Lócrio2TsTsTTTToolStripMenuItem.Click, AlgerianaTsTsssTsToolStripMenuItem.Click, OctatônicaTsTsTsTsTsToolStripMenuItem.Click, NonatônicaTssTsssTsToolStripMenuItem.Click, JaponesaAs2TTs2TToolStripMenuItem.Click, Skriabin2sTTTTToolStripMenuItem.Click, Javanesa1sTTT2TToolStripMenuItem.Click, Israelita1sTsTTsTsToolStripMenuItem.Click, JudaicaAdonaiMalakhsssTTTsTToolStripMenuItem.Click
        Try

            Dim NomeMenu As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
            VGoo = NomeMenu.Text
            NomeEscala = NomeMenu.Text
            Escalas.GeraEscalas()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Teclado_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        Try

            If ToolTipNota IsNot Nothing Then e.Graphics.DrawImage(ToolTipNota, 479, 74, 104, 58)
            e.Graphics.FillRectangle(Brushes.Yellow, OitavaEscala, 136, 238, 1)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

End Class