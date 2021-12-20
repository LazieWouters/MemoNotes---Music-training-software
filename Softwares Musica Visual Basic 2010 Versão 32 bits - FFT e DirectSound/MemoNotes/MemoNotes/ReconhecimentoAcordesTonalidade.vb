Option Strict Off
Option Explicit On

Public Class ReconhecimentoAcordesTonalidade

    Dim QtdeAcordes, ContadorAcorde(1), Acorde, ValorTickTimer, Tonalidade, spin, NotaMidi(3), GrauAcordeA(10), GrauAcordeB As Integer
    Dim NomeAcorde(100), Progressão As String
    'Dim Img As Bitmap = My.Resources.ReconhecimentoAcordesTonalidade
    Dim CorretoIncorreto As Bitmap
    Dim Cor1 As SolidBrush = New SolidBrush(Color.FromArgb(50, 0, 0, 255))
    Dim Tonalidades(,) As String = {{"C", "Dm", "Em", "F", "G7", "Am", "Bm7(b5)"}, _
{"G", "Am", "Bm", "C", "D7", "Em", "F#m7(b5)"}, _
{"D", "Em", "F#m", "G", "A7", "Bm", "C#m7(b5)"}, _
{"A", "Bm", "C#m", "D", "E7", "F#m", "G#m7(b5)"}, _
{"E", "F#m", "G#m", "A", "B7", "C#m", "D#m7(b5)"}, _
{"B", "C#m", "D#m", "E", "F#7", "G#m", "A#m7(b5)"}, _
{"Cb", "Dbm", "Ebm", "Fb", "Gb7", "Abm", "Bbm7(b5)"}, _
{"F#", "G#m", "A#m", "B", "C#7", "D#m", "E#m7(b5)"}, _
{"Gb", "Abm", "Bbm", "Cb", "Db7", "Ebm", "Fm7(b5)"}, _
{"C#", "D#m", "E#m", "F#", "G#7", "A#m", "B#m7(b5)"}, _
{"Db", "Ebm", "Fm", "Gb", "Ab7", "Bbm", "Cm7(b5)"}, _
{"Ab", "Bbm", "Cm", "Db", "Eb7", "Fm", "Gm7(b5)"}, _
{"Eb", "Fm", "Gm", "Ab", "Bb7", "Cm", "Dm7(b5)"}, _
{"Bb", "Cm", "Dm", "Eb", "F7", "Gm", "Am7(b5)"}, _
{"F", "Gm", "Am", "Bb", "C7", "Dm", "Em7(b5)"}}
    Dim Tonalidades2(,) As String = {{"C7M", "Dm7", "Em7", "F7M", "G7", "Am7", "Bm7(b5)"}, _
{"G7M", "Am7", "Bm7", "C7M", "D7", "Em7", "F#m7(b5)"}, _
{"D7M", "Em7", "F#m7", "G7M", "A7", "Bm7", "C#m7(b5)"}, _
{"A7M", "Bm7", "C#m7", "D7M", "E7", "F#m7", "G#m7(b5)"}, _
{"E7M", "F#m7", "G#m7", "A7M", "B7", "C#m7", "D#m7(b5)"}, _
{"B7M", "C#m7", "D#m7", "E7M", "F#7", "G#m7", "A#m7(b5)"}, _
{"Cb7M", "Dbm7", "Ebm7", "Fb7M", "Gb7", "Abm7", "Bbm7(b5)"}, _
{"F#7M", "G#m7", "A#m7", "B7M", "C#7", "D#m7", "E#m7(b5)"}, _
{"Gb7M", "Abm7", "Bbm7", "Cb7M", "Db7", "Ebm7", "Fm7(b5)"}, _
{"C#7M", "D#m7", "E#m7", "F#7M", "G#7", "A#m7", "B#m7(b5)"}, _
{"Db7M", "Ebm7", "Fm7", "Gb7M", "Ab7", "Bbm7", "Cm7(b5)"}, _
{"Ab7M", "Bbm7", "Cm7", "Db7M", "Eb7", "Fm7", "Gm7(b5)"}, _
{"Eb7M", "Fm7", "Gm7", "Ab7M", "Bb7", "Cm7", "Dm7(b5)"}, _
{"Bb7M", "Cm7", "Dm7", "Eb7M", "F7", "Gm7", "Am7(b5)"}, _
{"F7M", "Gm7", "Am7", "Bb7M", "C7", "Dm7", "Em7(b5)"}}
    Dim TonalidadesArray3D(,,) As Integer = _
{{{60, 64, 67, 71}, {62, 65, 69, 72}, {64, 67, 71, 74}, {65, 69, 72, 76}, {67, 71, 74, 77}, {69, 72, 76, 79}, {71, 74, 77, 81}}, _
{{55, 59, 62, 66}, {57, 60, 64, 67}, {59, 62, 66, 69}, {60, 64, 67, 71}, {62, 66, 69, 72}, {64, 67, 71, 74}, {66, 69, 72, 76}}, _
{{62, 66, 69, 73}, {64, 67, 71, 74}, {66, 69, 73, 76}, {67, 71, 74, 78}, {69, 73, 76, 79}, {71, 74, 78, 81}, {73, 76, 79, 83}}, _
{{57, 61, 64, 68}, {59, 62, 66, 69}, {61, 64, 68, 71}, {62, 66, 69, 73}, {64, 68, 71, 74}, {66, 69, 73, 76}, {68, 71, 74, 78}}, _
{{64, 68, 71, 75}, {66, 69, 73, 76}, {68, 71, 75, 78}, {69, 73, 76, 80}, {71, 75, 78, 81}, {73, 76, 80, 83}, {75, 78, 81, 85}}, _
{{59, 63, 66, 70}, {61, 64, 68, 71}, {63, 66, 70, 73}, {64, 68, 71, 75}, {66, 70, 73, 76}, {68, 71, 75, 78}, {70, 73, 76, 80}}, _
{{59, 63, 66, 70}, {61, 64, 68, 71}, {63, 66, 70, 73}, {64, 68, 71, 75}, {66, 70, 73, 76}, {68, 71, 75, 78}, {70, 73, 76, 80}}, _
{{66, 70, 73, 77}, {68, 71, 75, 78}, {70, 73, 77, 80}, {71, 75, 78, 82}, {73, 77, 80, 83}, {75, 78, 82, 85}, {77, 80, 83, 87}}, _
{{66, 70, 73, 77}, {68, 71, 75, 78}, {70, 73, 77, 80}, {71, 75, 78, 82}, {73, 77, 80, 83}, {75, 78, 82, 85}, {77, 80, 83, 87}}, _
{{61, 65, 68, 72}, {63, 66, 70, 73}, {65, 68, 72, 75}, {66, 70, 73, 77}, {68, 72, 75, 78}, {70, 73, 77, 80}, {72, 75, 78, 82}}, _
{{61, 65, 68, 72}, {63, 66, 70, 73}, {65, 68, 72, 75}, {66, 70, 73, 77}, {68, 72, 75, 78}, {70, 73, 77, 80}, {72, 75, 78, 82}}, _
{{56, 60, 63, 67}, {58, 61, 65, 68}, {60, 63, 67, 70}, {61, 65, 68, 72}, {63, 67, 70, 73}, {65, 68, 72, 75}, {67, 70, 73, 77}}, _
{{63, 67, 70, 74}, {65, 68, 72, 75}, {67, 70, 74, 77}, {68, 72, 75, 79}, {70, 74, 77, 80}, {72, 75, 79, 82}, {74, 77, 80, 84}}, _
{{58, 62, 65, 69}, {60, 63, 67, 70}, {62, 65, 69, 72}, {63, 67, 70, 74}, {65, 69, 72, 75}, {67, 70, 74, 77}, {69, 72, 75, 79}}, _
{{65, 69, 72, 76}, {67, 70, 74, 77}, {69, 72, 76, 79}, {70, 74, 77, 81}, {72, 76, 79, 82}, {74, 77, 81, 84}, {76, 79, 82, 86}}}


    Private Sub ReconhecimentoAcordesTonalidade_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
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

    Private Sub ReconhecimentoAcordesTonalidade_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If CorretoIncorreto IsNot Nothing Then e.Graphics.DrawImage(CorretoIncorreto, 185, 142, 349, 61)
    End Sub

    Private Sub SalvaSettings()
        Try

            My.Settings.NovoValorTonalidades(0) = TonalidadeC.Checked
            My.Settings.NovoValorTonalidades(1) = TonalidadeD.Checked
            My.Settings.NovoValorTonalidades(2) = TonalidadeE.Checked
            My.Settings.NovoValorTonalidades(3) = TonalidadeF.Checked
            My.Settings.NovoValorTonalidades(4) = TonalidadeG.Checked
            My.Settings.NovoValorTonalidades(5) = TonalidadeA.Checked
            My.Settings.NovoValorTonalidades(6) = TonalidadeB.Checked
            My.Settings.NovoValorTonalidades(7) = TonalidadeCb.Checked
            My.Settings.NovoValorTonalidades(8) = TonalidadeDb.Checked
            My.Settings.NovoValorTonalidades(9) = TonalidadeEb.Checked
            My.Settings.NovoValorTonalidades(10) = TonalidadeGb.Checked
            My.Settings.NovoValorTonalidades(11) = TonalidadeAb.Checked
            My.Settings.NovoValorTonalidades(12) = TonalidadeBb.Checked
            My.Settings.NovoValorTonalidades(13) = TonalidadeCsust.Checked
            My.Settings.NovoValorTonalidades(14) = TonalidadeFsust.Checked
            My.Settings.NovoValorQuantidadeDeAcordes = QuantidadeDeAcordes.Text
            My.Settings.NovoValorTimers(6) = Minutos.Text
            My.Settings.NovoValorTimers(7) = Segundos.Text
            My.Settings.NovoValorTimers(8) = Decimos.Text
            My.Settings.NovoValorTonicaComoUltimoAcorde = CheckBox1.Checked
            My.Settings.NovoValorInstrumentoMusical(3) = ComboBox1.Text

            My.Settings.NovoValorGraus(0) = CheckBox2.Checked
            My.Settings.NovoValorGraus(1) = CheckBox3.Checked
            My.Settings.NovoValorGraus(2) = CheckBox4.Checked
            My.Settings.NovoValorGraus(3) = CheckBox5.Checked
            My.Settings.NovoValorGraus(4) = CheckBox6.Checked
            My.Settings.NovoValorGraus(5) = CheckBox7.Checked
            My.Settings.NovoValorTocarTetrade = TocarTétrade.Checked

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub ReconhecimentoAcordesTonalidade_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            NotaMidi(0) = 1000
            NotaMidi(1) = 1000
            NotaMidi(2) = 1000
            NotaMidi(3) = 1000

            MidiPlayer.OpenMidi()
        Catch ex As Exception
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

            'The color at Pixel(10,10) is rendered as transparent for the complete background. 
            'Img.MakeTransparent(Img.GetPixel(10, 10))
            'Me.BackgroundImage = Img
            'Me.TransparencyKey = Img.GetPixel(10, 10)
            Me.BringToFront()


            TonalidadeC.Checked = My.Settings.NovoValorTonalidades(0)
            TonalidadeD.Checked = My.Settings.NovoValorTonalidades(1)
            TonalidadeE.Checked = My.Settings.NovoValorTonalidades(2)
            TonalidadeF.Checked = My.Settings.NovoValorTonalidades(3)
            TonalidadeG.Checked = My.Settings.NovoValorTonalidades(4)
            TonalidadeA.Checked = My.Settings.NovoValorTonalidades(5)
            TonalidadeB.Checked = My.Settings.NovoValorTonalidades(6)
            TonalidadeCb.Checked = My.Settings.NovoValorTonalidades(7)
            TonalidadeDb.Checked = My.Settings.NovoValorTonalidades(8)
            TonalidadeEb.Checked = My.Settings.NovoValorTonalidades(9)
            TonalidadeGb.Checked = My.Settings.NovoValorTonalidades(10)
            TonalidadeAb.Checked = My.Settings.NovoValorTonalidades(11)
            TonalidadeBb.Checked = My.Settings.NovoValorTonalidades(12)
            TonalidadeCsust.Checked = My.Settings.NovoValorTonalidades(13)
            TonalidadeFsust.Checked = My.Settings.NovoValorTonalidades(14)
            QuantidadeDeAcordes.Text = My.Settings.NovoValorQuantidadeDeAcordes
            Minutos.Text = My.Settings.NovoValorTimers(6)
            Segundos.Text = My.Settings.NovoValorTimers(7)
            Decimos.Text = My.Settings.NovoValorTimers(8)
            CheckBox1.Checked = My.Settings.NovoValorTonicaComoUltimoAcorde

            CheckBox2.Checked = My.Settings.NovoValorGraus(0)
            CheckBox3.Checked = My.Settings.NovoValorGraus(1)
            CheckBox4.Checked = My.Settings.NovoValorGraus(2)
            CheckBox5.Checked = My.Settings.NovoValorGraus(3)
            CheckBox6.Checked = My.Settings.NovoValorGraus(4)
            CheckBox7.Checked = My.Settings.NovoValorGraus(5)
            TocarTétrade.Checked = My.Settings.NovoValorTocarTetrade


            ComboBox1.Items.Add("Acordes de Violão")
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

            ComboBox1.Text = My.Settings.NovoValorInstrumentoMusical(3)

            If ComboBox1.Text = "" Then ComboBox1.Text = "000 - Acoustic Grand Piano"

            Tocanota = 1
            GeraProgressão()
        End Try

    End Sub

    Private Sub GeraProgressão() Handles Button9.Click, TocarTétrade.CheckedChanged

        Try

            Timer1.Enabled = False
            Button8.Enabled = True
            CorretoIncorreto = Nothing

            Dim Rect1 As New Rectangle(185, 142, 349, 61)
            Me.Invalidate(Rect1)

            Tonalidade = 500
            ContadorAcorde(0) = 1
            ContadorAcorde(1) = 1
            QtdeAcordes = CInt(QuantidadeDeAcordes.Text)
            Label3.Text = ""

            Do While Tonalidade > 14

                ' Create a byte array to hold the random value.
                Dim randomNumber(0) As Byte

                ' Create a new instance of the RNGCryptoServiceProvider.
                Dim Gen As New Security.Cryptography.RNGCryptoServiceProvider()

                ' Fill the array with a random value.
                Gen.GetBytes(randomNumber)

                ' Convert the byte to an integer value to make the modulus operation easier.
                Tonalidade = Int(Convert.ToInt32(randomNumber(0)))
                If (Tonalidade = 0 AndAlso Not TonalidadeC.Checked) OrElse
                (Tonalidade = 1 AndAlso Not TonalidadeG.Checked) OrElse
                (Tonalidade = 2 AndAlso Not TonalidadeD.Checked) OrElse
                (Tonalidade = 3 AndAlso Not TonalidadeA.Checked) OrElse
                (Tonalidade = 4 AndAlso Not TonalidadeE.Checked) OrElse
                (Tonalidade = 5 AndAlso Not TonalidadeB.Checked) OrElse
                (Tonalidade = 6 AndAlso Not TonalidadeCb.Checked) OrElse
                (Tonalidade = 7 AndAlso Not TonalidadeFsust.Checked) OrElse
                (Tonalidade = 8 AndAlso Not TonalidadeGb.Checked) OrElse
                (Tonalidade = 9 AndAlso Not TonalidadeCsust.Checked) OrElse
                (Tonalidade = 10 AndAlso Not TonalidadeDb.Checked) OrElse
                (Tonalidade = 11 AndAlso Not TonalidadeAb.Checked) OrElse
                (Tonalidade = 12 AndAlso Not TonalidadeEb.Checked) OrElse
                (Tonalidade = 13 AndAlso Not TonalidadeBb.Checked) OrElse
                (Tonalidade = 14 AndAlso Not TonalidadeF.Checked) Then
                    Tonalidade = 500
                End If

            Loop

            NomeBotões()

            Do While ContadorAcorde(0) <= QtdeAcordes

                Acorde = 500
                Do While Acorde > 6

                    ' Create a byte array to hold the random value.
                    Dim randomNumber(0) As Byte

                    ' Create a new instance of the RNGCryptoServiceProvider.
                    Dim Gen As New Security.Cryptography.RNGCryptoServiceProvider()

                    ' Fill the array with a random value.
                    Gen.GetBytes(randomNumber)

                    ' Convert the byte to an integer value to make the modulus operation easier.
                    Acorde = Int(Convert.ToInt32(randomNumber(0)))

                    If Acorde = 1 AndAlso Not CheckBox2.Checked Then
                        Acorde = 500
                    ElseIf Acorde = 2 AndAlso Not CheckBox3.Checked Then
                        Acorde = 500
                    ElseIf Acorde = 3 AndAlso Not CheckBox4.Checked Then
                        Acorde = 500
                    ElseIf Acorde = 4 AndAlso Not CheckBox5.Checked Then
                        Acorde = 500
                    ElseIf Acorde = 5 AndAlso Not CheckBox6.Checked Then
                        Acorde = 500
                    ElseIf Acorde = 6 AndAlso Not CheckBox7.Checked Then
                        Acorde = 500
                    End If



                Loop

                If TocarTétrade.Checked = False Then
                    NomeAcorde(ContadorAcorde(0)) = Tonalidades(Tonalidade, Acorde)
                Else
                    NomeAcorde(ContadorAcorde(0)) = Tonalidades2(Tonalidade, Acorde)
                End If

                GrauAcordeA(ContadorAcorde(0)) = Acorde

                If NomeAcorde(ContadorAcorde(0)) = NomeAcorde(ContadorAcorde(0) - 1) Then ContadorAcorde(0) -= 1
                ContadorAcorde(0) += 1


            Loop


            If TocarTétrade.Checked = False Then
                NomeAcorde(1) = Tonalidades(Tonalidade, 0) 'primeiro acorde será sempre a tônica
            Else
                NomeAcorde(1) = Tonalidades2(Tonalidade, 0) 'primeiro acorde será sempre a tônica
            End If

            GrauAcordeA(1) = 0 'primeiro acorde será sempre a tônica

            If CheckBox1.Checked Then
                If TocarTétrade.Checked = False Then
                    NomeAcorde(QtdeAcordes) = Tonalidades(Tonalidade, 0) 'último acorde será sempre a tônica
                Else
                    NomeAcorde(QtdeAcordes) = Tonalidades2(Tonalidade, 0) 'último acorde será sempre a tônica
                End If
                GrauAcordeA(QtdeAcordes) = 0 'último acorde será sempre a tônica
            End If


            ValorTickTimer = 1
            Label4.Text = ""
            Progressão = ""
            Timer1.Interval = 1
            Timer1.Enabled = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub NomeBotões()
        Try

            If TocarTétrade.Checked = False Then
                Label2.Text = Tonalidades(Tonalidade, 0)
                Button1.Text = Tonalidades(Tonalidade, 0)
                Button2.Text = Tonalidades(Tonalidade, 1)
                Button3.Text = Tonalidades(Tonalidade, 2)
                Button4.Text = Tonalidades(Tonalidade, 3)
                Button5.Text = Tonalidades(Tonalidade, 4)
                Button6.Text = Tonalidades(Tonalidade, 5)
                Button7.Text = Tonalidades(Tonalidade, 6)
            Else
                Label2.Text = Tonalidades2(Tonalidade, 0)
                Button1.Text = Tonalidades2(Tonalidade, 0)
                Button2.Text = Tonalidades2(Tonalidade, 1)
                Button3.Text = Tonalidades2(Tonalidade, 2)
                Button4.Text = Tonalidades2(Tonalidade, 3)
                Button5.Text = Tonalidades2(Tonalidade, 4)
                Button6.Text = Tonalidades2(Tonalidade, 5)
                Button7.Text = Tonalidades2(Tonalidade, 6)
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

            If ValorTickTimer <= QtdeAcordes Then
                Progressão = Progressão & NomeAcorde(ValorTickTimer) & "   "
                Label4.Text = Label4.Text & "?   "
                NomeAcorde(100) = NomeAcorde(ValorTickTimer)
                GrauAcordeB = GrauAcordeA(ValorTickTimer)

                Sons()

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

    Private Sub Button_Click(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Button7.MouseDown, Button6.MouseDown, Button5.MouseDown, Button4.MouseDown, Button3.MouseDown, Button2.MouseDown, Button1.MouseDown

        Try

            ' Obtém referência ao botão que invocou este método.
            Dim pbutton As Button = DirectCast(sender, Button)


            NomeAcorde(100) = pbutton.Text

            If pbutton.Name = "Button1" Then
                GrauAcordeB = 0
            ElseIf pbutton.Name = "Button2" Then
                GrauAcordeB = 1
            ElseIf pbutton.Name = "Button3" Then
                GrauAcordeB = 2
            ElseIf pbutton.Name = "Button4" Then
                GrauAcordeB = 3
            ElseIf pbutton.Name = "Button5" Then
                GrauAcordeB = 4
            ElseIf pbutton.Name = "Button6" Then
                GrauAcordeB = 5
            ElseIf pbutton.Name = "Button7" Then
                GrauAcordeB = 6
            End If


            If e.Button = MouseButtons.Left Then

                If ContadorAcorde(1) <= QtdeAcordes Then
                    Label3.Text = Label3.Text & NomeAcorde(100) & "   "
                    Sons()
                End If

                If ContadorAcorde(1) = QtdeAcordes Then
                    Button8.Enabled = False
                    If Label3.Text = Progressão Then
                        CorretoIncorreto = My.Resources.Correto
                        Timer2.Enabled = True
                    Else
                        CorretoIncorreto = My.Resources.Incorreto
                    End If

                    Dim Rect1 As New Rectangle(185, 142, 349, 61)
                    Me.Invalidate(Rect1)

                    Label4.Text = Progressão


                End If

                ContadorAcorde(1) += 1

            Else
                Sons()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub Sons()

        Try
            Try
                If NotaMidi(0) <> 1000 Then MidiPlayer.Play(New NoteOff(0, 1, CByte(NotaMidi(0)), CByte(NumericUpDown2.Value)))
                If NotaMidi(1) <> 1000 Then MidiPlayer.Play(New NoteOff(0, 1, CByte(NotaMidi(1)), CByte(NumericUpDown2.Value)))
                If NotaMidi(2) <> 1000 Then MidiPlayer.Play(New NoteOff(0, 1, CByte(NotaMidi(2)), CByte(NumericUpDown2.Value)))
                If NotaMidi(3) <> 1000 Then MidiPlayer.Play(New NoteOff(0, 1, CByte(NotaMidi(3)), CByte(NumericUpDown2.Value)))
            Catch ex As Exception

                'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
            Finally
                'esta parte do bloco é executada independentemente se algum erro acontecer ou não

            End Try

            If ComboBox1.Text = "Acordes de Violão" Then
                AjustaAcordesEnarmonicos()
                mp3 = "Sons\Acordes_Violao\" & NomeAcorde(100) & ".mp3"
                'TocaSom()
            Else
                Timer3.Enabled = False
                Try
                    MidiPlayer.Play(New ProgramChange(0, 1, CType(ComboBox1.Text.Substring(0, 3), GeneralMidiInstruments)))
                    MidiPlayer.Play(New NoteOn(0, 1, CByte(TonalidadesArray3D(Tonalidade, GrauAcordeB, 0)), CByte(NumericUpDown2.Value)))
                    MidiPlayer.Play(New NoteOn(0, 1, CByte(TonalidadesArray3D(Tonalidade, GrauAcordeB, 1)), CByte(NumericUpDown2.Value)))
                    MidiPlayer.Play(New NoteOn(0, 1, CByte(TonalidadesArray3D(Tonalidade, GrauAcordeB, 2)), CByte(NumericUpDown2.Value)))
                    If TocarTétrade.Checked = True Then MidiPlayer.Play(New NoteOn(0, 1, CByte(TonalidadesArray3D(Tonalidade, GrauAcordeB, 3)), CByte(NumericUpDown2.Value)))
                Catch ex As Exception

                    'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
                Finally
                    'esta parte do bloco é executada independentemente se algum erro acontecer ou não

                End Try

                NotaMidi(0) = TonalidadesArray3D(Tonalidade, GrauAcordeB, 0)
                NotaMidi(1) = TonalidadesArray3D(Tonalidade, GrauAcordeB, 1)
                NotaMidi(2) = TonalidadesArray3D(Tonalidade, GrauAcordeB, 2)
                NotaMidi(3) = TonalidadesArray3D(Tonalidade, GrauAcordeB, 3)

                Timer3.Enabled = True
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub AjustaAcordesEnarmonicos()

        Try

            NomeAcorde(100) = NomeAcorde(100).Replace("Cb", "B")
            NomeAcorde(100) = NomeAcorde(100).Replace("Db", "C#")
            NomeAcorde(100) = NomeAcorde(100).Replace("Eb", "D#")
            NomeAcorde(100) = NomeAcorde(100).Replace("Fb", "E")
            NomeAcorde(100) = NomeAcorde(100).Replace("Gb", "F#")
            NomeAcorde(100) = NomeAcorde(100).Replace("Ab", "G#")
            NomeAcorde(100) = NomeAcorde(100).Replace("Bb", "A#")
            NomeAcorde(100) = NomeAcorde(100).Replace("E#", "F")
            NomeAcorde(100) = NomeAcorde(100).Replace("B#", "C")


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        ValorTickTimer = 1
        Label4.Text = ""
        Progressão = ""
        Timer1.Interval = 1
        Timer1.Enabled = True
    End Sub

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click
        If CInt(QuantidadeDeAcordes.Text) < 10 Then QuantidadeDeAcordes.Text = CStr(CDbl(QuantidadeDeAcordes.Text) + 1)
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        If CInt(QuantidadeDeAcordes.Text) > 1 Then QuantidadeDeAcordes.Text = CStr(CDbl(QuantidadeDeAcordes.Text) - 1)
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
            Timer1.Interval = CInt(((CDbl(Minutos.Text) * 60) + CDbl(Segundos.Text) + (CDbl(Decimos.Text) / 100)) * 1000) : Timer3.Interval = Timer1.Interval


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

            Timer1.Interval = CInt(((CDbl(Minutos.Text) * 60) + CDbl(Segundos.Text) + (CDbl(Decimos.Text) / 100)) * 1000) : Timer3.Interval = Timer1.Interval

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
            Timer1.Interval = CInt(((CDbl(Minutos.Text) * 60) + CDbl(Segundos.Text) + (CDbl(Decimos.Text) / 100)) * 1000) : Timer3.Interval = Timer1.Interval

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
            Timer1.Interval = CInt(((CDbl(Minutos.Text) * 60) + CDbl(Segundos.Text) + (CDbl(Decimos.Text) / 100)) * 1000) : Timer3.Interval = Timer1.Interval

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
            Timer1.Interval = CInt(((CDbl(Minutos.Text) * 60) + CDbl(Segundos.Text) + (CDbl(Decimos.Text) / 100)) * 1000) : Timer3.Interval = Timer1.Interval

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub ab_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown, PictureBox27.MouseDown, PictureBox26.MouseDown, PictureBox2.MouseDown, PictureBox1.MouseDown, Label2.MouseDown, TonalidadeDb.MouseDown, TonalidadeCb.MouseDown, TonalidadeB.MouseDown, TonalidadeA.MouseDown, TonalidadeG.MouseDown, TonalidadeF.MouseDown, TonalidadeE.MouseDown, TonalidadeD.MouseDown, TonalidadeFsust.MouseDown, TonalidadeCsust.MouseDown, TonalidadeBb.MouseDown, TonalidadeAb.MouseDown, TonalidadeGb.MouseDown, TonalidadeEb.MouseDown, TonalidadeC.MouseDown, PictureBoxGb.MouseDown, PictureBoxG.MouseDown, PictureBoxFsust.MouseDown, PictureBoxF.MouseDown, PictureBoxEb.MouseDown, PictureBoxE.MouseDown, PictureBoxDb.MouseDown, PictureBoxD.MouseDown, PictureBoxCsust.MouseDown, PictureBoxCb.MouseDown, PictureBoxC.MouseDown, PictureBoxBb.MouseDown, PictureBoxB.MouseDown, PictureBoxAb.MouseDown, PictureBoxA.MouseDown, Label4.MouseDown, Label3.MouseDown
        VGa = Me.MousePosition.X - Me.Location.X
        VGb = Me.MousePosition.Y - Me.Location.Y
    End Sub

    Private Sub ab_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove, Segundos.MouseMove, PictureBox27.MouseMove, PictureBox26.MouseMove, PictureBox2.MouseMove, PictureBox1.MouseMove, Minutos.MouseMove, QuantidadeDeAcordes.MouseMove, Label2.MouseMove, Decimos.MouseMove, TonalidadeDb.MouseMove, TonalidadeCb.MouseMove, TonalidadeB.MouseMove, TonalidadeA.MouseMove, TonalidadeG.MouseMove, TonalidadeF.MouseMove, TonalidadeE.MouseMove, TonalidadeD.MouseMove, TonalidadeFsust.MouseMove, TonalidadeCsust.MouseMove, TonalidadeBb.MouseMove, TonalidadeAb.MouseMove, TonalidadeGb.MouseMove, TonalidadeEb.MouseMove, TonalidadeC.MouseMove, PictureBoxGb.MouseMove, PictureBoxG.MouseMove, PictureBoxFsust.MouseMove, PictureBoxF.MouseMove, PictureBoxEb.MouseMove, PictureBoxE.MouseMove, PictureBoxDb.MouseMove, PictureBoxD.MouseMove, PictureBoxCsust.MouseMove, PictureBoxCb.MouseMove, PictureBoxC.MouseMove, PictureBoxBb.MouseMove, PictureBoxB.MouseMove, PictureBoxAb.MouseMove, PictureBoxA.MouseMove, Label4.MouseMove, Label3.MouseMove
        If e.Button = MouseButtons.Left Then
            VGnewPoint = Me.MousePosition
            VGnewPoint.X -= VGa
            VGnewPoint.Y -= VGb
            Me.Location = VGnewPoint
        End If
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        GeraProgressão()
        Timer2.Enabled = False
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
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

    Private Sub PictureBoxC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBoxC.Click
        TonalidadeC.Checked = Not TonalidadeC.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBoxD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBoxD.Click
        TonalidadeD.Checked = Not TonalidadeD.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBoxE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBoxE.Click
        TonalidadeE.Checked = Not TonalidadeE.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBoxF_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBoxF.Click
        TonalidadeF.Checked = Not TonalidadeF.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBoxG_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBoxG.Click
        TonalidadeG.Checked = Not TonalidadeG.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBoxA_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBoxA.Click
        TonalidadeA.Checked = Not TonalidadeA.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBoxB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBoxB.Click
        TonalidadeB.Checked = Not TonalidadeB.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBoxCb_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBoxCb.Click
        TonalidadeCb.Checked = Not TonalidadeCb.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBoxDb_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBoxDb.Click
        TonalidadeDb.Checked = Not TonalidadeDb.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBoxEb_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBoxEb.Click
        TonalidadeEb.Checked = Not TonalidadeEb.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBoxGb_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBoxGb.Click
        TonalidadeGb.Checked = Not TonalidadeGb.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBoxAb_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBoxAb.Click
        TonalidadeAb.Checked = Not TonalidadeAb.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBoxBb_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBoxBb.Click
        TonalidadeBb.Checked = Not TonalidadeBb.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBoxCsust_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBoxCsust.Click
        TonalidadeCsust.Checked = Not TonalidadeCsust.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBoxFsust_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBoxFsust.Click
        TonalidadeFsust.Checked = Not TonalidadeFsust.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBoxC_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBoxC.Paint
        If TonalidadeC.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBoxD_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBoxD.Paint
        If TonalidadeD.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBoxE_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBoxE.Paint
        If TonalidadeE.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBoxF_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBoxF.Paint
        If TonalidadeF.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBoxG_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBoxG.Paint
        If TonalidadeG.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBoxA_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBoxA.Paint
        If TonalidadeA.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBoxB_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBoxB.Paint
        If TonalidadeB.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBoxCb_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBoxCb.Paint
        If TonalidadeCb.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBoxDb_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBoxDb.Paint
        If TonalidadeDb.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBoxEb_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBoxEb.Paint
        If TonalidadeEb.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBoxGb_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBoxGb.Paint
        If TonalidadeGb.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBoxAb_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBoxAb.Paint
        If TonalidadeAb.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBoxBb_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBoxBb.Paint
        If TonalidadeBb.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBoxCsust_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBoxCsust.Paint
        If TonalidadeCsust.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBoxFsust_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBoxFsust.Paint
        If TonalidadeFsust.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub ChecarFigurasNecessárias()

        Try

            If Not TonalidadeC.Checked AndAlso Not TonalidadeD.Checked AndAlso _
            Not TonalidadeE.Checked AndAlso Not TonalidadeF.Checked AndAlso _
            Not TonalidadeG.Checked AndAlso Not TonalidadeA.Checked AndAlso _
            Not TonalidadeB.Checked AndAlso Not TonalidadeCb.Checked AndAlso _
            Not TonalidadeDb.Checked AndAlso Not TonalidadeEb.Checked AndAlso _
            Not TonalidadeGb.Checked AndAlso Not TonalidadeAb.Checked AndAlso _
            Not TonalidadeBb.Checked AndAlso Not TonalidadeCsust.Checked AndAlso _
            Not TonalidadeFsust.Checked Then TonalidadeC.Checked = True : PictureBoxC.Invalidate()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub CancelaSom()

        Try
            MidiPlayer.Play(New NoteOff(0, 1, CByte(NotaMidi(0)), CByte(NumericUpDown2.Value)))
            MidiPlayer.Play(New NoteOff(0, 1, CByte(NotaMidi(1)), CByte(NumericUpDown2.Value)))
            MidiPlayer.Play(New NoteOff(0, 1, CByte(NotaMidi(2)), CByte(NumericUpDown2.Value)))
            MidiPlayer.Play(New NoteOff(0, 1, CByte(NotaMidi(3)), CByte(NumericUpDown2.Value)))
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

    Private Sub TocarTétrade_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TocarTétrade.CheckedChanged, ComboBox1.TextChanged
        If ComboBox1.Text = "Acordes de Violão" Then TocarTétrade.Checked = False
        NomeBotões()
    End Sub
End Class