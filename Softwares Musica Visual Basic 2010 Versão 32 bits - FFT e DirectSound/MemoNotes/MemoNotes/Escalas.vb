Option Strict On
Option Explicit On

Public Class Escalas
    Inherits PerPixelAlphaForm 'para mecher no form no modo design é preciso mudar para 'Inherits Form'
    'e depois de ajustado voltar para 'Inherits PerPixelAlphaForm' e clicar para mudar o inherit
    Public TransAmount As Byte = 255

    Dim aaa, bbb, eee, ppp, bolinhaIntervalo As Integer
    Dim scales As String
    Dim ddd As PictureBox
    Private DecimaThread As Thread
    Dim CorBolinha(21) As SolidBrush
    Dim CorrigeFundoDóSustenido(21) As Integer
    Dim posicaoScales(1, 1) As Integer
    Dim Topo(,) As Integer = {{195, 179, 179, 195, 179, 179, 195, 195, 195, 195, 179, 179, 195, 179, 179, 195, 179, 179, 195, 195, 195, 195}, _
{195, 179, 179, 195, 179, 179, 195, 195, 195, 195, 179, 179, 195, 179, 179, 195, 179, 179, 195, 195, 195, 195}, _
{549, 565, 565, 549, 565, 565, 565, 549, 565, 549, 565, 565, 549, 565, 565, 549, 565, 565, 565, 549, 565, 549}, _
{549, 565, 565, 549, 565, 565, 565, 549, 565, 549, 565, 565, 549, 565, 565, 549, 565, 565, 565, 549, 565, 549}, _
{549, 565, 565, 549, 565, 565, 565, 549, 565, 549, 565, 565, 549, 565, 565, 549, 565, 565, 565, 549, 565, 549}, _
{549, 565, 565, 549, 565, 565, 565, 549, 565, 549, 565, 565, 549, 565, 565, 549, 565, 565, 565, 549, 565, 549}, _
{343, 327, 327, 343, 343, 343, 327, 343, 327, 343, 327, 327, 343, 327, 327, 343, 343, 343, 327, 343, 327, 343}, _
{343, 327, 327, 343, 343, 343, 327, 343, 327, 343, 327, 327, 343, 327, 327, 343, 343, 343, 327, 343, 327, 343}, _
{401, 417, 417, 417, 401, 401, 417, 401, 417, 401, 417, 417, 401, 417, 417, 417, 401, 401, 417, 401, 417, 401}, _
{401, 417, 417, 417, 401, 401, 417, 401, 417, 401, 417, 417, 401, 417, 417, 417, 401, 401, 417, 401, 417, 401}, _
{401, 417, 417, 417, 401, 401, 417, 401, 417, 401, 417, 417, 401, 417, 417, 417, 401, 401, 417, 401, 417, 401}, _
{401, 417, 417, 417, 401, 401, 417, 401, 417, 401, 417, 417, 401, 417, 417, 417, 401, 401, 417, 401, 417, 401}, _
{491, 491, 491, 475, 491, 491, 475, 491, 475, 491, 475, 475, 491, 491, 491, 475, 491, 491, 475, 491, 475, 491}, _
{491, 491, 491, 475, 491, 491, 475, 491, 475, 491, 475, 475, 491, 491, 491, 475, 491, 491, 475, 491, 475, 491}, _
{269, 253, 253, 269, 253, 253, 269, 253, 269, 253, 269, 269, 269, 253, 253, 269, 253, 253, 269, 269, 269, 269}, _
{269, 253, 253, 269, 253, 253, 269, 253, 269, 253, 269, 269, 269, 253, 253, 269, 253, 253, 269, 269, 269, 269}, _
{623, 639, 639, 623, 639, 639, 623, 639, 623, 639, 639, 639, 623, 639, 639, 623, 639, 639, 639, 623, 639, 623}, _
{623, 639, 639, 623, 639, 639, 623, 639, 623, 639, 639, 639, 623, 639, 639, 623, 639, 639, 639, 623, 639, 623}, _
{623, 639, 639, 623, 639, 639, 623, 639, 623, 639, 639, 639, 623, 639, 639, 623, 639, 639, 639, 623, 639, 623}, _
{623, 639, 639, 623, 639, 639, 623, 639, 623, 639, 639, 639, 623, 639, 639, 623, 639, 639, 639, 623, 639, 623}, _
{269, 253, 253, 269, 253, 253, 269, 269, 269, 269, 253, 253, 269, 253, 253, 269, 269, 269, 253, 269, 253, 269}, _
{269, 253, 253, 269, 253, 253, 269, 269, 269, 269, 253, 253, 269, 253, 253, 269, 269, 269, 253, 269, 253, 269}, _
{475, 491, 491, 475, 491, 491, 491, 475, 491, 475, 491, 491, 475, 491, 491, 491, 475, 475, 491, 475, 491, 475}, _
{475, 491, 491, 475, 491, 491, 491, 475, 491, 475, 491, 491, 475, 491, 491, 491, 475, 475, 491, 475, 491, 475}, _
{475, 491, 491, 475, 491, 491, 491, 475, 491, 475, 491, 491, 475, 491, 491, 491, 475, 475, 491, 475, 491, 475}, _
{475, 491, 491, 475, 491, 491, 491, 475, 491, 475, 491, 491, 475, 491, 491, 491, 475, 475, 491, 475, 491, 475}, _
{417, 401, 401, 417, 417, 417, 401, 417, 401, 417, 401, 401, 417, 417, 417, 401, 417, 417, 401, 417, 401, 417}, _
{417, 401, 401, 417, 417, 417, 401, 417, 401, 417, 401, 401, 417, 417, 417, 401, 417, 417, 401, 417, 401, 417}, _
{328, 344, 344, 344, 328, 328, 344, 328, 344, 328, 344, 344, 344, 328, 328, 344, 328, 328, 344, 328, 344, 328}, _
{328, 344, 344, 344, 328, 328, 344, 328, 344, 328, 344, 344, 344, 328, 328, 344, 328, 328, 344, 328, 344, 328}, _
{328, 344, 344, 344, 328, 328, 344, 328, 344, 328, 344, 344, 344, 328, 328, 344, 328, 328, 344, 328, 344, 328}, _
{328, 344, 344, 344, 328, 328, 344, 328, 344, 328, 344, 344, 344, 328, 328, 344, 328, 328, 344, 328, 344, 328}, _
{565, 565, 565, 549, 565, 565, 549, 565, 549, 565, 565, 565, 549, 565, 565, 549, 565, 565, 549, 565, 549, 565}, _
{565, 565, 565, 549, 565, 565, 549, 565, 549, 565, 565, 565, 549, 565, 565, 549, 565, 565, 549, 565, 549, 565}}
    Dim Esquerda(,) As Integer = {{367, 372, 372, 376, 381, 381, 385, 394, 385, 394, 399, 399, 403, 408, 408, 412, 417, 417, 421, 430, 421, 430}, _
{367, 372, 372, 376, 381, 381, 385, 394, 385, 394, 399, 399, 403, 408, 408, 412, 417, 417, 421, 430, 421, 430}, _
{174, 178, 178, 183, 187, 187, 196, 201, 196, 201, 205, 205, 210, 214, 214, 219, 223, 223, 232, 237, 232, 237}, _
{174, 178, 178, 183, 187, 187, 196, 201, 196, 201, 205, 205, 210, 214, 214, 219, 223, 223, 232, 237, 232, 237}, _
{174, 178, 178, 183, 187, 187, 196, 201, 196, 201, 205, 205, 210, 214, 214, 219, 223, 223, 232, 237, 232, 237}, _
{174, 178, 178, 183, 187, 187, 196, 201, 196, 201, 205, 205, 210, 214, 214, 219, 223, 223, 232, 237, 232, 237}, _
{639, 644, 644, 648, 657, 657, 662, 666, 662, 666, 671, 671, 675, 680, 680, 684, 693, 693, 698, 702, 698, 702}, _
{639, 644, 644, 648, 657, 657, 662, 666, 662, 666, 671, 671, 675, 680, 680, 684, 693, 693, 698, 702, 698, 702}, _
{80, 84, 84, 93, 98, 98, 102, 107, 102, 107, 111, 111, 116, 120, 120, 129, 134, 134, 138, 143, 138, 143}, _
{80, 84, 84, 93, 98, 98, 102, 107, 102, 107, 111, 111, 116, 120, 120, 129, 134, 134, 138, 143, 138, 143}, _
{80, 84, 84, 93, 98, 98, 102, 107, 102, 107, 111, 111, 116, 120, 120, 129, 134, 134, 138, 143, 138, 143}, _
{80, 84, 84, 93, 98, 98, 102, 107, 102, 107, 111, 111, 116, 120, 120, 129, 134, 134, 138, 143, 138, 143}, _
{648, 657, 657, 662, 666, 666, 671, 675, 671, 675, 680, 680, 684, 693, 693, 698, 702, 702, 707, 711, 707, 711}, _
{648, 657, 657, 662, 666, 666, 671, 675, 671, 675, 680, 680, 684, 693, 693, 698, 702, 702, 707, 711, 707, 711}, _
{196, 201, 201, 205, 210, 210, 214, 219, 214, 219, 223, 223, 232, 237, 237, 241, 246, 246, 250, 259, 250, 259}, _
{196, 201, 201, 205, 210, 210, 214, 219, 214, 219, 223, 223, 232, 237, 237, 241, 246, 246, 250, 259, 250, 259}, _
{399, 403, 403, 408, 412, 412, 417, 421, 417, 421, 430, 430, 435, 439, 439, 444, 448, 448, 457, 462, 457, 462}, _
{399, 403, 403, 408, 412, 412, 417, 421, 417, 421, 430, 430, 435, 439, 439, 444, 448, 448, 457, 462, 457, 462}, _
{399, 403, 403, 408, 412, 412, 417, 421, 417, 421, 430, 430, 435, 439, 439, 444, 448, 448, 457, 462, 457, 462}, _
{399, 403, 403, 408, 412, 412, 417, 421, 417, 421, 430, 430, 435, 439, 439, 444, 448, 448, 457, 462, 457, 462}, _
{603, 608, 608, 612, 617, 617, 621, 630, 621, 630, 635, 635, 639, 644, 644, 648, 657, 657, 662, 666, 662, 666}, _
{603, 608, 608, 612, 617, 617, 621, 630, 621, 630, 635, 635, 639, 644, 644, 648, 657, 657, 662, 666, 662, 666}, _
{147, 151, 151, 156, 160, 160, 169, 174, 169, 174, 178, 178, 183, 187, 187, 196, 201, 201, 205, 210, 205, 210}, _
{147, 151, 151, 156, 160, 160, 169, 174, 169, 174, 178, 178, 183, 187, 187, 196, 201, 201, 205, 210, 205, 210}, _
{147, 151, 151, 156, 160, 160, 169, 174, 169, 174, 178, 178, 183, 187, 187, 196, 201, 201, 205, 210, 205, 210}, _
{147, 151, 151, 156, 160, 160, 169, 174, 169, 174, 178, 178, 183, 187, 187, 196, 201, 201, 205, 210, 205, 210}, _
{714, 719, 719, 723, 732, 732, 737, 741, 737, 741, 746, 746, 750, 759, 759, 764, 768, 768, 773, 777, 773, 777}, _
{714, 719, 719, 723, 732, 732, 737, 741, 737, 741, 746, 746, 750, 759, 759, 764, 768, 768, 773, 777, 773, 777}, _
{156, 160, 160, 169, 174, 174, 178, 183, 178, 183, 187, 187, 196, 201, 201, 205, 210, 210, 214, 219, 214, 219}, _
{156, 160, 160, 169, 174, 174, 178, 183, 178, 183, 187, 187, 196, 201, 201, 205, 210, 210, 214, 219, 214, 219}, _
{156, 160, 160, 169, 174, 174, 178, 183, 178, 183, 187, 187, 196, 201, 201, 205, 210, 210, 214, 219, 214, 219}, _
{156, 160, 160, 169, 174, 174, 178, 183, 178, 183, 187, 187, 196, 201, 201, 205, 210, 210, 214, 219, 214, 219}, _
{621, 630, 630, 635, 639, 639, 644, 648, 644, 648, 657, 657, 662, 666, 666, 671, 675, 675, 680, 684, 680, 684}, _
{621, 630, 630, 635, 639, 639, 644, 648, 644, 648, 657, 657, 662, 666, 666, 671, 675, 675, 680, 684, 680, 684}}
    Dim ArrayIntervalo(,) As String = {{"C", "C#", "Db", "D", "D#", "Eb", "E", "E#", "Fb", "F", "F#", "Gb", "G", "G#", "Ab", "A", "A#", "Bb", "B", "B#", "Cb", "C"}, _
{"C", "C#", "Db", "D", "D#", "Eb", "E", "F", "E", "F", "F#", "Gb", "G", "G#", "Ab", "A", "A#", "Bb", "B", "C", "B", "C"}, _
{"C#", "C##", "D", "D#", "D##", "Fb", "F", "F#", "F", "F#", "F##", "G", "G#", "G##", "A", "A#", "A##", "Cb", "C", "C#", "C", "C#"}, _
{"C#", "D", "D", "D#", "E", "E", "F", "F#", "F", "F#", "G", "G", "G#", "A", "A", "A#", "B", "B", "C", "C#", "C", "C#"}, _
{"Db", "D", "Ebb", "Eb", "E", "Fb", "F", "F#", "Gbb", "Gb", "G", "Abb", "Ab", "A", "Bbb", "Bb", "B", "Cb", "C", "C#", "Dbb", "Db"}, _
{"Db", "D", "D", "Eb", "E", "E", "F", "F#", "F", "Gb", "G", "G", "Ab", "A", "A", "Bb", "B", "B", "C", "C#", "C", "Db"}, _
{"D", "D#", "Eb", "E", "E#", "F", "F#", "F##", "Gb", "G", "G#", "Ab", "A", "A#", "Bb", "B", "B#", "C", "C#", "C##", "Db", "D"}, _
{"D", "D#", "Eb", "E", "F", "F", "F#", "G", "Gb", "G", "G#", "Ab", "A", "A#", "Bb", "B", "C", "C", "C#", "D", "Db", "D"}, _
{"D#", "D##", "Fb", "F", "F#", "Gb", "G", "G#", "G", "G#", "G##", "A", "A#", "A##", "Cb", "C", "C#", "Db", "D", "D#", "D", "D#"}, _
{"D#", "E", "E", "F", "F#", "Gb", "G", "G#", "G", "G#", "A", "A", "A#", "B", "B", "C", "C#", "Db", "D", "D#", "D", "D#"}, _
{"Eb", "E", "Fb", "F", "F#", "Gb", "G", "G#", "Abb", "Ab", "A", "Bbb", "Bb", "B", "Cb", "C", "C#", "Db", "D", "D#", "Ebb", "Eb"}, _
{"Eb", "E", "E", "F", "F#", "Gb", "G", "G#", "G", "Ab", "A", "A", "Bb", "B", "B", "C", "C#", "Db", "D", "D#", "D", "Eb"}, _
{"E", "E#", "F", "F#", "F##", "G", "G#", "G##", "Ab", "A", "A#", "Bb", "B", "B#", "C", "C#", "C##", "D", "D#", "D##", "Eb", "E"}, _
{"E", "F", "F", "F#", "G", "G", "G#", "A", "Ab", "A", "A#", "Bb", "B", "C", "C", "C#", "D", "D", "D#", "E", "Eb", "E"}, _
{"F", "F#", "Gb", "G", "G#", "Ab", "A", "A#", "Bbb", "Bb", "B", "Cb", "C", "C#", "Db", "D", "D#", "Eb", "E", "E#", "Fb", "F"}, _
{"F", "F#", "Gb", "G", "G#", "Ab", "A", "A#", "A", "Bb", "B", "B", "C", "C#", "Db", "D", "D#", "Eb", "E", "E#", "E", "F"}, _
{"F#", "F##", "G", "G#", "G##", "A", "A#", "A##", "Bb", "B", "B#", "C", "C#", "C##", "D", "D#", "D##", "Fb", "F", "F#", "F", "F#"}, _
{"F#", "G", "G", "G#", "A", "A", "A#", "B", "Bb", "B", "C", "C", "C#", "D", "D", "D#", "E", "E", "F", "F#", "F", "F#"}, _
{"Gb", "G", "Abb", "Ab", "A", "Bbb", "Bb", "B", "Bb", "B", "B#", "Dbb", "Db", "D", "Ebb", "Eb", "E", "Fb", "F", "F#", "Gbb", "Gb"}, _
{"Gb", "G", "G", "Ab", "A", "A", "Bb", "B", "Bb", "B", "C", "C", "Db", "D", "D", "Eb", "E", "E", "F", "F#", "F", "Gb"}, _
{"G", "G#", "Ab", "A", "A#", "Bb", "B", "B#", "Cb", "C", "C#", "Db", "D", "D#", "Eb", "E", "E#", "F", "F#", "F##", "Gb", "G"}, _
{"G", "G#", "Ab", "A", "A#", "Bb", "B", "C", "Cb", "C", "C#", "Db", "D", "D#", "Eb", "E", "F", "F", "F#", "G", "Gb", "G"}, _
{"G#", "G##", "A", "A#", "A##", "Cb", "C", "C#", "C", "C#", "C##", "D", "D#", "D##", "Fb", "F", "F#", "Gb", "G", "G#", "G", "G#"}, _
{"G#", "A", "A", "A#", "B", "B", "C", "C#", "C", "C#", "D", "D", "D#", "E", "E", "F", "F#", "Gb", "G", "G#", "G", "G#"}, _
{"Ab", "A", "Bbb", "Bb", "B", "Cb", "C", "C#", "Dbb", "Db", "D", "Ebb", "Eb", "E", "Fb", "F", "F#", "Gb", "G", "G#", "Abb", "Ab"}, _
{"Ab", "A", "A", "Bb", "B", "B", "C", "C#", "C", "Db", "D", "D", "Eb", "E", "E", "F", "F#", "Gb", "G", "G#", "G", "Ab"}, _
{"A", "A#", "Bb", "B", "B#", "C", "C#", "C##", "Db", "D", "D#", "Eb", "E", "E#", "F", "F#", "F##", "G", "G#", "G##", "Ab", "A"}, _
{"A", "A#", "Bb", "B", "C", "C", "C#", "D", "Db", "D", "D#", "Eb", "E", "F", "F", "F#", "G", "G", "G#", "A", "Ab", "A"}, _
{"A#", "A##", "Cb", "C", "C#", "Db", "D", "D#", "D", "D#", "D##", "Fb", "F", "F#", "Gb", "G", "G#", "Ab", "A", "A#", "A", "A#"}, _
{"A#", "B", "B", "C", "C#", "Db", "D", "D#", "D", "D#", "E", "E", "F", "F#", "Gb", "G", "G#", "Ab", "A", "A#", "A", "A#"}, _
{"Bb", "B", "Cb", "C", "C#", "Db", "D", "D#", "Ebb", "Eb", "E", "Fb", "F", "F#", "Gb", "G", "G#", "Ab", "A", "A#", "Bbb", "Bb"}, _
{"Bb", "B", "B", "C", "C#", "Db", "D", "D#", "D", "Eb", "E", "E", "F", "F#", "Gb", "G", "G#", "Ab", "A", "A#", "A", "Bb"}, _
{"B", "B#", "C", "C#", "C##", "D", "D#", "D##", "Eb", "E", "E#", "F", "F#", "F##", "G", "G#", "G##", "A", "A#", "A##", "Bb", "B"}, _
{"B", "C", "C", "C#", "D", "D", "D#", "E", "Eb", "E", "F", "F", "F#", "G", "G", "G#", "A", "A", "A#", "B", "Bb", "B"}}

    Public Sub exibeControles() Handles Me.Load

        Try
            Dim gr As Graphics
            Dim FaceBit As New Bitmap(My.Resources.Escalas)
            gr = Graphics.FromImage(FaceBit)
            gr.SmoothingMode = SmoothingMode.AntiAlias

            aaa = 0

            Dim Preto As SolidBrush = New SolidBrush(ColorDialogC.Color)
            Dim Verde As SolidBrush = New SolidBrush(ColorDialogD.Color)
            Dim Azul As SolidBrush = New SolidBrush(ColorDialogE.Color)
            Dim Vermelho As SolidBrush = New SolidBrush(ColorDialogF.Color)
            Dim Amarelo As SolidBrush = New SolidBrush(ColorDialogG.Color)
            Dim Laranja As SolidBrush = New SolidBrush(ColorDialogA.Color)
            Dim Rosa As SolidBrush = New SolidBrush(ColorDialogB.Color)
            Dim Branco As SolidBrush = New SolidBrush(Color.FromArgb(255, 255, 250))
            Dim CorRetangulo As Pen = New Pen(Color.FromArgb(2, 0, 0), 1)
            Dim myFontScale As New Font("Arial", 11, FontStyle.Bold, GraphicsUnit.Pixel)
            Dim Branco2 As Pen = New Pen(Color.FromArgb(255, 255, 250), 1)


            Dim Fonte As New Font("Arial", 18, FontStyle.Regular, GraphicsUnit.Pixel)
            Dim MedidasTexto As SizeF = gr.MeasureString(NomeEscala, Fonte)
            Dim MedidasTexto2 As SizeF = gr.MeasureString(Intervalos, myFontScale)
            Dim PosicaoX As Integer = CInt(((Me.Width + 9) / 2) - (MedidasTexto.Width / 2))
            Dim PosicaoY As Integer = 88
            Dim PosicaoX2 As Integer = CInt(((Me.Width + 9) / 2) - (MedidasTexto2.Width / 2))
            Dim PosicaoY2 As Integer = 110

            gr.DrawString(NomeEscala, Fonte, Branco, PosicaoX, PosicaoY)
            gr.DrawString(Intervalos, myFontScale, Branco, PosicaoX2, PosicaoY2)


            If NomeEscala <> "" Then
                Do While aaa <= 32


                    If Not Teclado.CheckBox200.Checked AndAlso (aaa = 2 OrElse aaa = 8 OrElse aaa = 16 OrElse aaa = 22 OrElse aaa = 28) Then
                        aaa += 2
                    ElseIf Teclado.CheckBox200.Checked AndAlso (aaa = 4 OrElse aaa = 10 OrElse aaa = 18 OrElse aaa = 24 OrElse aaa = 30) Then
                        aaa += 2
                    End If


                    If Teclado.Simplificado.Checked Then bbb = 1 Else bbb = 0

                    Dim largura, altura As Integer
                    largura = 6
                    altura = 6
                    eee = 0
                    Do While eee <= 21

                        If Teclado.CheckBox100.Checked Then
                            If ArrayIntervalo(aaa + bbb, eee) = "C" OrElse ArrayIntervalo(aaa + bbb, eee) = "Cb" OrElse ArrayIntervalo(aaa + bbb, eee) = "C##" Then
                                CorBolinha(eee) = Preto
                                CorrigeFundoDóSustenido(eee) = 0
                            ElseIf ArrayIntervalo(aaa + bbb, eee) = "C#" Then
                                CorBolinha(eee) = Preto
                                If Preto.Color = (Color.FromArgb(2, 0, 0)) Then
                                    CorrigeFundoDóSustenido(eee) = 1
                                Else
                                    CorrigeFundoDóSustenido(eee) = 0
                                End If
                            ElseIf ArrayIntervalo(aaa + bbb, eee) = "D" OrElse ArrayIntervalo(aaa + bbb, eee) = "D#" OrElse ArrayIntervalo(aaa + bbb, eee) = "D##" OrElse ArrayIntervalo(aaa + bbb, eee) = "Db" OrElse ArrayIntervalo(aaa + bbb, eee) = "Dbb" Then
                                CorBolinha(eee) = Verde
                                CorrigeFundoDóSustenido(eee) = 0
                            ElseIf ArrayIntervalo(aaa + bbb, eee) = "E" OrElse ArrayIntervalo(aaa + bbb, eee) = "Eb" OrElse ArrayIntervalo(aaa + bbb, eee) = "Ebb" OrElse ArrayIntervalo(aaa + bbb, eee) = "E#" Then
                                CorBolinha(eee) = Azul
                                CorrigeFundoDóSustenido(eee) = 0
                            ElseIf ArrayIntervalo(aaa + bbb, eee) = "F" OrElse ArrayIntervalo(aaa + bbb, eee) = "Fb" OrElse ArrayIntervalo(aaa + bbb, eee) = "F#" OrElse ArrayIntervalo(aaa + bbb, eee) = "F##" Then
                                CorBolinha(eee) = Vermelho
                                CorrigeFundoDóSustenido(eee) = 0
                            ElseIf ArrayIntervalo(aaa + bbb, eee) = "G" OrElse ArrayIntervalo(aaa + bbb, eee) = "Gb" OrElse ArrayIntervalo(aaa + bbb, eee) = "Gbb" OrElse ArrayIntervalo(aaa + bbb, eee) = "G#" OrElse ArrayIntervalo(aaa + bbb, eee) = "G##" Then
                                CorBolinha(eee) = Amarelo
                                CorrigeFundoDóSustenido(eee) = 0
                            ElseIf ArrayIntervalo(aaa + bbb, eee) = "A" OrElse ArrayIntervalo(aaa + bbb, eee) = "Ab" OrElse ArrayIntervalo(aaa + bbb, eee) = "Abb" OrElse ArrayIntervalo(aaa + bbb, eee) = "A#" OrElse ArrayIntervalo(aaa + bbb, eee) = "A##" Then
                                CorBolinha(eee) = Laranja
                                CorrigeFundoDóSustenido(eee) = 0
                            ElseIf ArrayIntervalo(aaa + bbb, eee) = "B" OrElse ArrayIntervalo(aaa + bbb, eee) = "Bb" OrElse ArrayIntervalo(aaa + bbb, eee) = "Bbb" OrElse ArrayIntervalo(aaa + bbb, eee) = "B#" Then
                                CorBolinha(eee) = Rosa
                                CorrigeFundoDóSustenido(eee) = 0
                            End If
                        Else
                            CorBolinha(eee) = Vermelho
                            CorrigeFundoDóSustenido(eee) = 0
                        End If

                        eee += 1
                    Loop

                    scales = ""


                    If I_1_0 = 0 Then
                        gr.FillEllipse(CorBolinha(0), Esquerda(aaa + bbb, 0), Topo(aaa + bbb, 0), largura, altura) : scales = scales & ArrayIntervalo(aaa + bbb, I_1_0) & " "
                        If CorrigeFundoDóSustenido(0) = 1 Then gr.DrawEllipse(Branco2, Esquerda(aaa + bbb, 0) + 1, Topo(aaa + bbb, 0) + 1, largura - 2, altura - 2)
                    End If
                    If I_a1_1 = 1 Then
                        gr.FillEllipse(CorBolinha(1), Esquerda(aaa + bbb, 1), Topo(aaa + bbb, 1), largura, altura) : scales = scales & ArrayIntervalo(aaa + bbb, I_a1_1) & " "
                        If CorrigeFundoDóSustenido(1) = 1 Then gr.DrawEllipse(Branco2, Esquerda(aaa + bbb, 1) + 1, Topo(aaa + bbb, 1) + 1, largura - 2, altura - 2)
                    End If
                    If I_b2_2 = 2 Then
                        gr.FillEllipse(CorBolinha(2), Esquerda(aaa + bbb, 2), Topo(aaa + bbb, 2), largura, altura) : scales = scales & ArrayIntervalo(aaa + bbb, I_b2_2) & " "
                        If CorrigeFundoDóSustenido(2) = 1 Then gr.DrawEllipse(Branco2, Esquerda(aaa + bbb, 2) + 1, Topo(aaa + bbb, 2) + 1, largura - 2, altura - 2)
                    End If
                    If I_2_3 = 3 Then
                        gr.FillEllipse(CorBolinha(3), Esquerda(aaa + bbb, 3), Topo(aaa + bbb, 3), largura, altura) : scales = scales & ArrayIntervalo(aaa + bbb, I_2_3) & " "
                        If CorrigeFundoDóSustenido(3) = 1 Then gr.DrawEllipse(Branco2, Esquerda(aaa + bbb, 3) + 1, Topo(aaa + bbb, 3) + 1, largura - 2, altura - 2)
                    End If
                    If I_a2_4 = 4 Then
                        gr.FillEllipse(CorBolinha(4), Esquerda(aaa + bbb, 4), Topo(aaa + bbb, 4), largura, altura) : scales = scales & ArrayIntervalo(aaa + bbb, I_a2_4) & " "
                        If CorrigeFundoDóSustenido(4) = 1 Then gr.DrawEllipse(Branco2, Esquerda(aaa + bbb, 4) + 1, Topo(aaa + bbb, 4) + 1, largura - 2, altura - 2)
                    End If
                    If I_b3_5 = 5 Then
                        gr.FillEllipse(CorBolinha(5), Esquerda(aaa + bbb, 5), Topo(aaa + bbb, 5), largura, altura) : scales = scales & ArrayIntervalo(aaa + bbb, I_b3_5) & " "
                        If CorrigeFundoDóSustenido(5) = 1 Then gr.DrawEllipse(Branco2, Esquerda(aaa + bbb, 5) + 1, Topo(aaa + bbb, 5) + 1, largura - 2, altura - 2)
                    End If
                    If I_3_6 = 6 Then
                        gr.FillEllipse(CorBolinha(6), Esquerda(aaa + bbb, 6), Topo(aaa + bbb, 6), largura, altura) : scales = scales & ArrayIntervalo(aaa + bbb, I_3_6) & " "
                        If CorrigeFundoDóSustenido(6) = 1 Then gr.DrawEllipse(Branco2, Esquerda(aaa + bbb, 6) + 1, Topo(aaa + bbb, 6) + 1, largura - 2, altura - 2)
                    End If
                    If I_a3_7 = 7 Then
                        gr.FillEllipse(CorBolinha(7), Esquerda(aaa + bbb, 7), Topo(aaa + bbb, 7), largura, altura) : scales = scales & ArrayIntervalo(aaa + bbb, I_a3_7) & " "
                        If CorrigeFundoDóSustenido(7) = 1 Then gr.DrawEllipse(Branco2, Esquerda(aaa + bbb, 7) + 1, Topo(aaa + bbb, 7) + 1, largura - 2, altura - 2)
                    End If
                    If I_b4_8 = 8 Then
                        gr.FillEllipse(CorBolinha(8), Esquerda(aaa + bbb, 8), Topo(aaa + bbb, 8), largura, altura) : scales = scales & ArrayIntervalo(aaa + bbb, I_b4_8) & " "
                        If CorrigeFundoDóSustenido(8) = 1 Then gr.DrawEllipse(Branco2, Esquerda(aaa + bbb, 8) + 1, Topo(aaa + bbb, 8) + 1, largura - 2, altura - 2)
                    End If
                    If I_4_9 = 9 Then
                        gr.FillEllipse(CorBolinha(9), Esquerda(aaa + bbb, 9), Topo(aaa + bbb, 9), largura, altura) : scales = scales & ArrayIntervalo(aaa + bbb, I_4_9) & " "
                        If CorrigeFundoDóSustenido(9) = 1 Then gr.DrawEllipse(Branco2, Esquerda(aaa + bbb, 9) + 1, Topo(aaa + bbb, 9) + 1, largura - 2, altura - 2)
                    End If
                    If I_a4_10 = 10 Then
                        gr.FillEllipse(CorBolinha(10), Esquerda(aaa + bbb, 10), Topo(aaa + bbb, 10), largura, altura) : scales = scales & ArrayIntervalo(aaa + bbb, I_a4_10) & " "
                        If CorrigeFundoDóSustenido(10) = 1 Then gr.DrawEllipse(Branco2, Esquerda(aaa + bbb, 10) + 1, Topo(aaa + bbb, 10) + 1, largura - 2, altura - 2)
                    End If
                    If I_b5_11 = 11 Then
                        gr.FillEllipse(CorBolinha(11), Esquerda(aaa + bbb, 11), Topo(aaa + bbb, 11), largura, altura) : scales = scales & ArrayIntervalo(aaa + bbb, I_b5_11) & " "
                        If CorrigeFundoDóSustenido(11) = 1 Then gr.DrawEllipse(Branco2, Esquerda(aaa + bbb, 11) + 1, Topo(aaa + bbb, 11) + 1, largura - 2, altura - 2)
                    End If
                    If I_5_12 = 12 Then
                        gr.FillEllipse(CorBolinha(12), Esquerda(aaa + bbb, 12), Topo(aaa + bbb, 12), largura, altura) : scales = scales & ArrayIntervalo(aaa + bbb, I_5_12) & " "
                        If CorrigeFundoDóSustenido(12) = 1 Then gr.DrawEllipse(Branco2, Esquerda(aaa + bbb, 12) + 1, Topo(aaa + bbb, 12) + 1, largura - 2, altura - 2)
                    End If
                    If I_a5_13 = 13 Then
                        gr.FillEllipse(CorBolinha(13), Esquerda(aaa + bbb, 13), Topo(aaa + bbb, 13), largura, altura) : scales = scales & ArrayIntervalo(aaa + bbb, I_a5_13) & " "
                        If CorrigeFundoDóSustenido(13) = 1 Then gr.DrawEllipse(Branco2, Esquerda(aaa + bbb, 13) + 1, Topo(aaa + bbb, 13) + 1, largura - 2, altura - 2)
                    End If
                    If I_b6_14 = 14 Then
                        gr.FillEllipse(CorBolinha(14), Esquerda(aaa + bbb, 14), Topo(aaa + bbb, 14), largura, altura) : scales = scales & ArrayIntervalo(aaa + bbb, I_b6_14) & " "
                        If CorrigeFundoDóSustenido(14) = 1 Then gr.DrawEllipse(Branco2, Esquerda(aaa + bbb, 14) + 1, Topo(aaa + bbb, 14) + 1, largura - 2, altura - 2)
                    End If
                    If I_6_15 = 15 Then
                        gr.FillEllipse(CorBolinha(15), Esquerda(aaa + bbb, 15), Topo(aaa + bbb, 15), largura, altura) : scales = scales & ArrayIntervalo(aaa + bbb, I_6_15) & " "
                        If CorrigeFundoDóSustenido(15) = 1 Then gr.DrawEllipse(Branco2, Esquerda(aaa + bbb, 15) + 1, Topo(aaa + bbb, 15) + 1, largura - 2, altura - 2)
                    End If
                    If I_a6_16 = 16 Then
                        gr.FillEllipse(CorBolinha(16), Esquerda(aaa + bbb, 16), Topo(aaa + bbb, 16), largura, altura) : scales = scales & ArrayIntervalo(aaa + bbb, I_a6_16) & " "
                        If CorrigeFundoDóSustenido(16) = 1 Then gr.DrawEllipse(Branco2, Esquerda(aaa + bbb, 16) + 1, Topo(aaa + bbb, 16) + 1, largura - 2, altura - 2)
                    End If
                    If I_b7_17 = 17 Then
                        gr.FillEllipse(CorBolinha(17), Esquerda(aaa + bbb, 17), Topo(aaa + bbb, 17), largura, altura) : scales = scales & ArrayIntervalo(aaa + bbb, I_b7_17) & " "
                        If CorrigeFundoDóSustenido(17) = 1 Then gr.DrawEllipse(Branco2, Esquerda(aaa + bbb, 17) + 1, Topo(aaa + bbb, 17) + 1, largura - 2, altura - 2)
                    End If
                    If I_7_18 = 18 Then
                        gr.FillEllipse(CorBolinha(18), Esquerda(aaa + bbb, 18), Topo(aaa + bbb, 18), largura, altura) : scales = scales & ArrayIntervalo(aaa + bbb, I_7_18) & " "
                        If CorrigeFundoDóSustenido(18) = 1 Then gr.DrawEllipse(Branco2, Esquerda(aaa + bbb, 18) + 1, Topo(aaa + bbb, 18) + 1, largura - 2, altura - 2)
                    End If
                    If I_a7_19 = 19 Then
                        gr.FillEllipse(CorBolinha(19), Esquerda(aaa + bbb, 19), Topo(aaa + bbb, 19), largura, altura) : scales = scales & ArrayIntervalo(aaa + bbb, I_a7_19) & " "
                        If CorrigeFundoDóSustenido(19) = 1 Then gr.DrawEllipse(Branco2, Esquerda(aaa + bbb, 19) + 1, Topo(aaa + bbb, 19) + 1, largura - 2, altura - 2)
                    End If
                    If I_b8_20 = 20 Then
                        gr.FillEllipse(CorBolinha(20), Esquerda(aaa + bbb, 20), Topo(aaa + bbb, 20), largura, altura) : scales = scales & ArrayIntervalo(aaa + bbb, I_b8_20) & " "
                        If CorrigeFundoDóSustenido(20) = 1 Then gr.DrawEllipse(Branco2, Esquerda(aaa + bbb, 20) + 1, Topo(aaa + bbb, 20) + 1, largura - 2, altura - 2)
                    End If
                    If I_8_21 = 21 Then
                        gr.FillEllipse(CorBolinha(21), Esquerda(aaa + bbb, 21), Topo(aaa + bbb, 21), largura, altura) : scales = scales & ArrayIntervalo(aaa + bbb, I_8_21) & " "
                        If CorrigeFundoDóSustenido(21) = 1 Then gr.DrawEllipse(Branco2, Esquerda(aaa + bbb, 21) + 1, Topo(aaa + bbb, 21) + 1, largura - 2, altura - 2)
                    End If


                    If aaa = 0 Then
                        posicaoScales(0, 0) = 495 : posicaoScales(0, 1) = 192
                    ElseIf aaa = 2 OrElse aaa = 4 Then
                        posicaoScales(0, 0) = 297 : posicaoScales(0, 1) = 562
                    ElseIf aaa = 6 Then
                        posicaoScales(0, 0) = 758 : posicaoScales(0, 1) = 340
                    ElseIf aaa = 8 OrElse aaa = 10 Then
                        posicaoScales(0, 0) = 194 : posicaoScales(0, 1) = 414
                    ElseIf aaa = 12 Then
                        posicaoScales(0, 0) = 758 : posicaoScales(0, 1) = 488
                    ElseIf aaa = 14 Then
                        posicaoScales(0, 0) = 297 : posicaoScales(0, 1) = 266
                    ElseIf aaa = 16 OrElse aaa = 18 Then
                        posicaoScales(0, 0) = 495 : posicaoScales(0, 1) = 636
                    ElseIf aaa = 20 Then
                        posicaoScales(0, 0) = 696 : posicaoScales(0, 1) = 266
                    ElseIf aaa = 22 OrElse aaa = 24 Then
                        posicaoScales(0, 0) = 234 : posicaoScales(0, 1) = 488
                    ElseIf aaa = 26 Then
                        posicaoScales(0, 0) = 797 : posicaoScales(0, 1) = 414
                    ElseIf aaa = 28 OrElse aaa = 30 Then
                        posicaoScales(0, 0) = 234 : posicaoScales(0, 1) = 340
                    ElseIf aaa = 32 Then
                        posicaoScales(0, 0) = 696 : posicaoScales(0, 1) = 562
                    End If
                    gr.DrawString(scales, myFontScale, Branco, posicaoScales(0, 0), posicaoScales(0, 1))

                    aaa += 2

                Loop
            End If

            gr.FillRectangle(Preto, 332, 176, 10, 10)
            gr.DrawRectangle(CorRetangulo, 332, 176, 10, 10)
            gr.FillRectangle(Vermelho, 133, 250, 10, 10)
            gr.DrawRectangle(CorRetangulo, 133, 250, 10, 10)
            gr.FillRectangle(Rosa, 70, 315, 10, 10)
            gr.DrawRectangle(CorRetangulo, 70, 315, 10, 10)
            gr.FillRectangle(Laranja, 70, 338, 10, 10)
            gr.DrawRectangle(CorRetangulo, 70, 338, 10, 10)
            gr.FillRectangle(Azul, 30, 388, 10, 10)
            gr.DrawRectangle(CorRetangulo, 30, 388, 10, 10)
            gr.FillRectangle(Verde, 30, 411, 10, 10)
            gr.DrawRectangle(CorRetangulo, 30, 411, 10, 10)
            gr.FillRectangle(Laranja, 70, 462, 10, 10)
            gr.DrawRectangle(CorRetangulo, 70, 462, 10, 10)
            gr.FillRectangle(Amarelo, 70, 484, 10, 10)
            gr.DrawRectangle(CorRetangulo, 70, 484, 10, 10)
            gr.FillRectangle(Verde, 133, 536, 10, 10)
            gr.DrawRectangle(CorRetangulo, 133, 536, 10, 10)
            gr.FillRectangle(Preto, 133, 558, 10, 10)
            gr.DrawRectangle(CorRetangulo, 133, 558, 10, 10)
            gr.FillRectangle(Vermelho, 332, 610, 10, 10)
            gr.DrawRectangle(CorRetangulo, 332, 610, 10, 10)
            gr.FillRectangle(Amarelo, 332, 632, 10, 10)
            gr.DrawRectangle(CorRetangulo, 332, 632, 10, 10)
            gr.FillRectangle(Rosa, 937, 545, 10, 10)
            gr.DrawRectangle(CorRetangulo, 937, 545, 10, 10)
            gr.FillRectangle(Azul, 998, 471, 10, 10)
            gr.DrawRectangle(CorRetangulo, 998, 471, 10, 10)
            gr.FillRectangle(Laranja, 1038, 397, 10, 10)
            gr.DrawRectangle(CorRetangulo, 1038, 397, 10, 10)
            gr.FillRectangle(Verde, 998, 323, 10, 10)
            gr.DrawRectangle(CorRetangulo, 998, 323, 10, 10)
            gr.FillRectangle(Amarelo, 937, 250, 10, 10)
            gr.DrawRectangle(CorRetangulo, 937, 250, 10, 10)

            'exibição linear dos intervalos
            bolinhaIntervalo = 0
            If VGoo <> "" Then
                Dim CorBolinhaIntervalos As SolidBrush = New SolidBrush(ColorDialog1.Color)
                gr.FillEllipse(CorBolinhaIntervalos, 24, 149, 6, 6)

                If VGc > 0 Then bolinhaIntervalo += (VGc * 20) : gr.FillEllipse(CorBolinhaIntervalos, 24 + bolinhaIntervalo, 149, 6, 6)
                If VGd > 0 Then bolinhaIntervalo += (VGd * 20) : gr.FillEllipse(CorBolinhaIntervalos, 24 + bolinhaIntervalo, 149, 6, 6)
                If VGee > 0 Then bolinhaIntervalo += (VGee * 20) : gr.FillEllipse(CorBolinhaIntervalos, 24 + bolinhaIntervalo, 149, 6, 6)
                If VGf > 0 Then bolinhaIntervalo += (VGf * 20) : gr.FillEllipse(CorBolinhaIntervalos, 24 + bolinhaIntervalo, 149, 6, 6)
                If VGg > 0 Then bolinhaIntervalo += (VGg * 20) : gr.FillEllipse(CorBolinhaIntervalos, 24 + bolinhaIntervalo, 149, 6, 6)
                If VGh > 0 Then bolinhaIntervalo += (VGh * 20) : gr.FillEllipse(CorBolinhaIntervalos, 24 + bolinhaIntervalo, 149, 6, 6)
                If VGi > 0 Then bolinhaIntervalo += (VGi * 20) : gr.FillEllipse(CorBolinhaIntervalos, 24 + bolinhaIntervalo, 149, 6, 6)
                If VGj > 0 Then bolinhaIntervalo += (VGj * 20) : gr.FillEllipse(CorBolinhaIntervalos, 24 + bolinhaIntervalo, 149, 6, 6)
                If VGk > 0 Then bolinhaIntervalo += (VGk * 20) : gr.FillEllipse(CorBolinhaIntervalos, 24 + bolinhaIntervalo, 149, 6, 6)
                If VGl > 0 Then bolinhaIntervalo += (VGl * 20) : gr.FillEllipse(CorBolinhaIntervalos, 24 + bolinhaIntervalo, 149, 6, 6)
                If VGm > 0 Then bolinhaIntervalo += (VGm * 20) : gr.FillEllipse(CorBolinhaIntervalos, 24 + bolinhaIntervalo, 149, 6, 6)
                If VGn > 0 Then bolinhaIntervalo += (VGn * 20) : gr.FillEllipse(CorBolinhaIntervalos, 24 + bolinhaIntervalo, 149, 6, 6)
                If VGt > 0 Then bolinhaIntervalo += (VGt * 20) : gr.FillEllipse(CorBolinhaIntervalos, 24 + bolinhaIntervalo, 149, 6, 6)
            End If

            'exibição circular dos intervalos
            bolinhaIntervalo = 0

            If VGoo <> "" Then
                Dim CorBolinhaIntervalos As SolidBrush = New SolidBrush(ColorDialog1.Color)
                gr.FillEllipse(CorBolinhaIntervalos, 891, 130, 6, 6)
                ppp = 1
                Do While ppp <= 12
                    If ppp = 1 AndAlso VGc > 0 Then
                        bolinhaIntervalo += (VGc * 20)
                    ElseIf ppp = 2 AndAlso VGd > 0 Then
                        bolinhaIntervalo += (VGd * 20)
                    ElseIf ppp = 3 AndAlso VGee > 0 Then
                        bolinhaIntervalo += (VGee * 20)
                    ElseIf ppp = 4 AndAlso VGf > 0 Then
                        bolinhaIntervalo += (VGf * 20)
                    ElseIf ppp = 5 AndAlso VGg > 0 Then
                        bolinhaIntervalo += (VGg * 20)
                    ElseIf ppp = 6 AndAlso VGh > 0 Then
                        bolinhaIntervalo += (VGh * 20)
                    ElseIf ppp = 7 AndAlso VGi > 0 Then
                        bolinhaIntervalo += (VGi * 20)
                    ElseIf ppp = 8 AndAlso VGj > 0 Then
                        bolinhaIntervalo += (VGj * 20)
                    ElseIf ppp = 9 AndAlso VGk > 0 Then
                        bolinhaIntervalo += (VGk * 20)
                    ElseIf ppp = 10 AndAlso VGl > 0 Then
                        bolinhaIntervalo += (VGl * 20)
                    ElseIf ppp = 11 AndAlso VGm > 0 Then
                        bolinhaIntervalo += (VGm * 20)
                    ElseIf ppp = 12 AndAlso VGn > 0 Then
                        bolinhaIntervalo += (VGn * 20)
                    ElseIf ppp = 13 AndAlso VGt > 0 Then
                        bolinhaIntervalo += (VGt * 20)
                    End If


                    ppp += 1
                    If bolinhaIntervalo = 20 Then
                        gr.FillEllipse(CorBolinhaIntervalos, 899, 103, 6, 6)
                    ElseIf bolinhaIntervalo = 40 Then
                        gr.FillEllipse(CorBolinhaIntervalos, 916, 86, 6, 6)
                    ElseIf bolinhaIntervalo = 60 Then
                        gr.FillEllipse(CorBolinhaIntervalos, 942, 80, 6, 6)
                    ElseIf bolinhaIntervalo = 80 Then
                        gr.FillEllipse(CorBolinhaIntervalos, 967, 86, 6, 6)
                    ElseIf bolinhaIntervalo = 100 Then
                        gr.FillEllipse(CorBolinhaIntervalos, 986, 105, 6, 6)
                    ElseIf bolinhaIntervalo = 120 Then
                        gr.FillEllipse(CorBolinhaIntervalos, 993, 130, 6, 6)
                    ElseIf bolinhaIntervalo = 140 Then
                        gr.FillEllipse(CorBolinhaIntervalos, 986, 156, 6, 6)
                    ElseIf bolinhaIntervalo = 160 Then
                        gr.FillEllipse(CorBolinhaIntervalos, 967, 175, 6, 6)
                    ElseIf bolinhaIntervalo = 180 Then
                        gr.FillEllipse(CorBolinhaIntervalos, 942, 182, 6, 6)
                    ElseIf bolinhaIntervalo = 200 Then
                        gr.FillEllipse(CorBolinhaIntervalos, 916, 175, 6, 6)
                    ElseIf bolinhaIntervalo = 220 Then
                        gr.FillEllipse(CorBolinhaIntervalos, 897, 156, 6, 6)
                    End If

                Loop
            End If

            'cor das bolinhas das escalas
            Dim corBolinhaIntervalosCorC As SolidBrush = New SolidBrush(ColorDialogC.Color)
            Dim corBolinhaIntervalosCorD As SolidBrush = New SolidBrush(ColorDialogD.Color)
            Dim corBolinhaIntervalosCorE As SolidBrush = New SolidBrush(ColorDialogE.Color)
            Dim corBolinhaIntervalosCorF As SolidBrush = New SolidBrush(ColorDialogF.Color)
            Dim corBolinhaIntervalosCorG As SolidBrush = New SolidBrush(ColorDialogG.Color)
            Dim corBolinhaIntervalosCorA As SolidBrush = New SolidBrush(ColorDialogA.Color)
            Dim corBolinhaIntervalosCorB As SolidBrush = New SolidBrush(ColorDialogB.Color)

            gr.FillRectangle(corBolinhaIntervalosCorC, CorC.Left, CorC.Top, CorC.Width, CorC.Height)
            gr.FillRectangle(corBolinhaIntervalosCorD, CorD.Left, CorD.Top, CorD.Width, CorD.Height)
            gr.FillRectangle(corBolinhaIntervalosCorE, CorE.Left, CorE.Top, CorE.Width, CorE.Height)
            gr.FillRectangle(corBolinhaIntervalosCorF, CorF.Left, CorF.Top, CorF.Width, CorF.Height)
            gr.FillRectangle(corBolinhaIntervalosCorG, CorG.Left, CorG.Top, CorG.Width, CorG.Height)
            gr.FillRectangle(corBolinhaIntervalosCorA, CorA.Left, CorA.Top, CorA.Width, CorA.Height)
            gr.FillRectangle(corBolinhaIntervalosCorB, CorB.Left, CorB.Top, CorB.Width, CorB.Height)

            gr.SmoothingMode = SmoothingMode.None

            If VGoo = "Maior (T-T-s-T-T-T-s)" OrElse VGoo = "Jônio (T-T-s-T-T-T-s)" Then
                gr.DrawImage(My.Resources.ClaveC, 616, 155, 88, 52) 'Clave de Dó
                gr.DrawImage(My.Resources.ClaveC, 879, 302, 88, 52) 'Clave de Ré
                gr.DrawImage(My.Resources.SustenidoArmaduraClave, 879 + 30, 302 + 10, 6, 16)
                gr.DrawImage(My.Resources.SustenidoArmaduraClave, 879 + 37, 302 + 18, 6, 16)
                gr.DrawImage(My.Resources.ClaveC, 879, 449, 88, 52) 'Clave de Mi
                gr.DrawImage(My.Resources.SustenidoArmaduraClave, 879 + 30, 449 + 10, 6, 16)
                gr.DrawImage(My.Resources.SustenidoArmaduraClave, 879 + 37, 449 + 18, 6, 16)
                gr.DrawImage(My.Resources.SustenidoArmaduraClave, 879 + 44, 449 + 7, 6, 16)
                gr.DrawImage(My.Resources.SustenidoArmaduraClave, 879 + 51, 449 + 15, 6, 16)
                gr.DrawImage(My.Resources.ClaveC, 417, 228, 88, 52) 'Clave de Fá
                gr.DrawImage(My.Resources.BemolArmaduraClave, 417 + 30, 228 + 17, 6, 16)
                gr.DrawImage(My.Resources.ClaveC, 816, 228, 88, 52) 'Clave de Sol
                gr.DrawImage(My.Resources.SustenidoArmaduraClave, 816 + 30, 228 + 10, 6, 16)
                gr.DrawImage(My.Resources.ClaveC, 918, 375, 88, 52) 'Clave de Lá
                gr.DrawImage(My.Resources.SustenidoArmaduraClave, 918 + 30, 375 + 10, 6, 16)
                gr.DrawImage(My.Resources.SustenidoArmaduraClave, 918 + 37, 375 + 18, 6, 16)
                gr.DrawImage(My.Resources.SustenidoArmaduraClave, 918 + 44, 375 + 7, 6, 16)
                gr.DrawImage(My.Resources.ClaveC, 816, 524, 88, 52) 'Clave de Si
                gr.DrawImage(My.Resources.SustenidoArmaduraClave, 816 + 30, 524 + 10, 6, 16)
                gr.DrawImage(My.Resources.SustenidoArmaduraClave, 816 + 37, 524 + 18, 6, 16)
                gr.DrawImage(My.Resources.SustenidoArmaduraClave, 816 + 44, 524 + 7, 6, 16)
                gr.DrawImage(My.Resources.SustenidoArmaduraClave, 816 + 51, 524 + 15, 6, 16)
                gr.DrawImage(My.Resources.SustenidoArmaduraClave, 816 + 58, 524 + 23, 6, 16)
                If Not Teclado.CheckBox200.Checked Then 'bemois
                    gr.DrawImage(My.Resources.ClaveC, 417, 524, 88, 52) 'Clave de Réb
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 417 + 30, 524 + 17, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 417 + 37, 524 + 10, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 417 + 44, 524 + 20, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 417 + 51, 524 + 12, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 417 + 58, 524 + 22, 6, 16)
                    gr.DrawImage(My.Resources.ClaveC, 314, 375, 88, 52) 'Clave de Mib
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 314 + 30, 375 + 17, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 314 + 37, 375 + 10, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 314 + 44, 375 + 20, 6, 16)
                    gr.DrawImage(My.Resources.ClaveC, 616, 598, 88, 52) 'Clave de Solb
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 616 + 30, 598 + 17, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 616 + 37, 598 + 10, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 616 + 44, 598 + 20, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 616 + 51, 598 + 12, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 616 + 58, 598 + 22, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 616 + 65, 598 + 15, 6, 16)
                    gr.DrawImage(My.Resources.ClaveC, 354, 449, 88, 52) 'Clave de Láb
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 354 + 30, 449 + 17, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 354 + 37, 449 + 10, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 354 + 44, 449 + 20, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 354 + 51, 449 + 12, 6, 16)
                    gr.DrawImage(My.Resources.ClaveC, 354, 302, 88, 52)  'Clave de Sib
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 354 + 30, 302 + 17, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 354 + 37, 302 + 10, 6, 16)
                Else 'sustenidos
                    gr.DrawImage(My.Resources.ClaveC, 417, 524, 88, 52) 'Clave de Dó#
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 417 + 30, 524 + 10, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 417 + 37, 524 + 18, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 417 + 44, 524 + 7, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 417 + 51, 524 + 15, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 417 + 58, 524 + 23, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 417 + 65, 524 + 13, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 417 + 72, 524 + 20, 6, 16)
                    gr.DrawImage(My.Resources.ClaveC, 616, 598, 88, 52) 'Clave de Fá#
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 616 + 30, 598 + 10, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 616 + 37, 598 + 18, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 616 + 44, 598 + 7, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 616 + 51, 598 + 15, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 616 + 58, 598 + 23, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 616 + 65, 598 + 13, 6, 16)
                End If
            ElseIf VGoo = "Menor Natural ou Menor Pura (T-s-T-T-s-T-T)" OrElse VGoo = "Eólio (T-s-T-T-s-T-T)" OrElse VGoo = "Menor Melódica - Descendente (T-s-T-T-s-T-T)" Then
                gr.DrawImage(My.Resources.ClaveC, 616, 155, 88, 52) 'Clave de Dó
                gr.DrawImage(My.Resources.BemolArmaduraClave, 616 + 30, 155 + 17, 6, 16)
                gr.DrawImage(My.Resources.BemolArmaduraClave, 616 + 37, 155 + 10, 6, 16)
                gr.DrawImage(My.Resources.BemolArmaduraClave, 616 + 44, 155 + 20, 6, 16)
                gr.DrawImage(My.Resources.ClaveC, 879, 302, 88, 52) 'Clave de Ré
                gr.DrawImage(My.Resources.BemolArmaduraClave, 879 + 30, 302 + 17, 6, 16)
                gr.DrawImage(My.Resources.ClaveC, 879, 449, 88, 52) 'Clave de Mi
                gr.DrawImage(My.Resources.SustenidoArmaduraClave, 879 + 30, 449 + 10, 6, 16)
                gr.DrawImage(My.Resources.ClaveC, 417, 228, 88, 52) 'Clave de Fá
                gr.DrawImage(My.Resources.BemolArmaduraClave, 417 + 30, 228 + 17, 6, 16)
                gr.DrawImage(My.Resources.BemolArmaduraClave, 417 + 37, 228 + 10, 6, 16)
                gr.DrawImage(My.Resources.BemolArmaduraClave, 417 + 44, 228 + 20, 6, 16)
                gr.DrawImage(My.Resources.BemolArmaduraClave, 417 + 51, 228 + 12, 6, 16)
                gr.DrawImage(My.Resources.ClaveC, 816, 228, 88, 52)  'Clave de Sol
                gr.DrawImage(My.Resources.BemolArmaduraClave, 816 + 30, 228 + 17, 6, 16)
                gr.DrawImage(My.Resources.BemolArmaduraClave, 816 + 37, 228 + 10, 6, 16)
                gr.DrawImage(My.Resources.ClaveC, 918, 375, 88, 52) 'Clave de Lá
                gr.DrawImage(My.Resources.ClaveC, 816, 524, 88, 52) 'Clave de Si
                gr.DrawImage(My.Resources.SustenidoArmaduraClave, 816 + 30, 524 + 10, 6, 16)
                gr.DrawImage(My.Resources.SustenidoArmaduraClave, 816 + 37, 524 + 18, 6, 16)
                If Not Teclado.CheckBox200.Checked Then 'bemois
                    gr.DrawImage(My.Resources.ClaveC, 314, 375, 88, 52) 'Clave de Mib
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 314 + 30, 375 + 17, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 314 + 37, 375 + 10, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 314 + 44, 375 + 20, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 314 + 51, 375 + 12, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 314 + 58, 375 + 22, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 314 + 65, 375 + 15, 6, 16)
                    gr.DrawImage(My.Resources.ClaveC, 354, 449, 88, 52) 'Clave de Láb
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 354 + 30, 449 + 17, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 354 + 37, 449 + 10, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 354 + 44, 449 + 20, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 354 + 51, 449 + 12, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 354 + 58, 449 + 22, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 354 + 65, 449 + 15, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 354 + 72, 449 + 25, 6, 16)
                    gr.DrawImage(My.Resources.ClaveC, 354, 302, 88, 52) 'Clave de Sib
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 354 + 30, 302 + 17, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 354 + 37, 302 + 10, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 354 + 44, 302 + 20, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 354 + 51, 302 + 12, 6, 16)
                    gr.DrawImage(My.Resources.BemolArmaduraClave, 354 + 58, 302 + 22, 6, 16)
                Else 'sustenidos
                    gr.DrawImage(My.Resources.ClaveC, 417, 524, 88, 52) 'Clave de Dó#
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 417 + 30, 524 + 10, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 417 + 37, 524 + 18, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 417 + 44, 524 + 7, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 417 + 51, 524 + 15, 6, 16)
                    gr.DrawImage(My.Resources.ClaveC, 314, 375, 88, 52) 'Clave de Ré#
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 314 + 30, 375 + 10, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 314 + 37, 375 + 18, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 314 + 44, 375 + 7, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 314 + 51, 375 + 15, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 314 + 58, 375 + 23, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 314 + 65, 375 + 13, 6, 16)
                    gr.DrawImage(My.Resources.ClaveC, 616, 598, 88, 52) 'Clave de Fá#
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 616 + 30, 598 + 10, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 616 + 37, 598 + 18, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 616 + 44, 598 + 7, 6, 16)
                    gr.DrawImage(My.Resources.ClaveC, 354, 449, 88, 52) 'Clave de Sol#
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 354 + 30, 449 + 10, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 354 + 37, 449 + 18, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 354 + 44, 449 + 7, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 354 + 51, 449 + 15, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 354 + 58, 449 + 23, 6, 16)
                    gr.DrawImage(My.Resources.ClaveC, 354, 302, 88, 52) 'Clave de Lá#
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 354 + 30, 302 + 10, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 354 + 37, 302 + 18, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 354 + 44, 302 + 7, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 354 + 51, 302 + 15, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 354 + 58, 302 + 23, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 354 + 65, 302 + 13, 6, 16)
                    gr.DrawImage(My.Resources.SustenidoArmaduraClave, 354 + 72, 302 + 20, 6, 16)
                End If
            End If

            'exibe os botões para tocar as progressões
            If VGoo = "Maior (T-T-s-T-T-T-s)" OrElse VGoo = "Jônio (T-T-s-T-T-T-s)" OrElse VGoo = "Menor Natural ou Menor Pura (T-s-T-T-s-T-T)" OrElse VGoo = "Eólio (T-s-T-T-s-T-T)" OrElse VGoo = "Menor Melódica - Descendente (T-s-T-T-s-T-T)" Then
                gr.DrawImage(My.Resources.ProgressãoAcordes, TocaAcordesEscalaDó.Left, TocaAcordesEscalaDó.Top, 61, 15)
                gr.DrawImage(My.Resources.ProgressãoAcordes, TocaAcordesEscalaSol.Left, TocaAcordesEscalaSol.Top, 61, 15)
                gr.DrawImage(My.Resources.ProgressãoAcordes, TocaAcordesEscalaRé.Left, TocaAcordesEscalaRé.Top, 61, 15)
                gr.DrawImage(My.Resources.ProgressãoAcordes, TocaAcordesEscalaLá.Left, TocaAcordesEscalaLá.Top, 61, 15)
                gr.DrawImage(My.Resources.ProgressãoAcordes, TocaAcordesEscalaMi.Left, TocaAcordesEscalaMi.Top, 61, 15)
                gr.DrawImage(My.Resources.ProgressãoAcordes, TocaAcordesEscalaSi.Left, TocaAcordesEscalaSi.Top, 61, 15)
                gr.DrawImage(My.Resources.ProgressãoAcordes, TocaAcordesEscalaFáSus.Left, TocaAcordesEscalaFáSus.Top, 61, 15)
                gr.DrawImage(My.Resources.ProgressãoAcordes, TocaAcordesEscalaDóSus.Left, TocaAcordesEscalaDóSus.Top, 61, 15)
                gr.DrawImage(My.Resources.ProgressãoAcordes, TocaAcordesEscalaSolSus.Left, TocaAcordesEscalaSolSus.Top, 61, 15)
                gr.DrawImage(My.Resources.ProgressãoAcordes, TocaAcordesEscalaRéSus.Left, TocaAcordesEscalaRéSus.Top, 61, 15)
                gr.DrawImage(My.Resources.ProgressãoAcordes, TocaAcordesEscalaLáSus.Left, TocaAcordesEscalaLáSus.Top, 61, 15)
                gr.DrawImage(My.Resources.ProgressãoAcordes, TocaAcordesEscalaFá.Left, TocaAcordesEscalaFá.Top, 61, 15)
                TocaAcordesEscalaDó.Visible = True
                TocaAcordesEscalaSol.Visible = True
                TocaAcordesEscalaRé.Visible = True
                TocaAcordesEscalaLá.Visible = True
                TocaAcordesEscalaMi.Visible = True
                TocaAcordesEscalaSi.Visible = True
                TocaAcordesEscalaFáSus.Visible = True
                TocaAcordesEscalaDóSus.Visible = True
                TocaAcordesEscalaSolSus.Visible = True
                TocaAcordesEscalaRéSus.Visible = True
                TocaAcordesEscalaLáSus.Visible = True
                TocaAcordesEscalaFá.Visible = True
            Else
                TocaAcordesEscalaDó.Visible = False
                TocaAcordesEscalaSol.Visible = False
                TocaAcordesEscalaRé.Visible = False
                TocaAcordesEscalaLá.Visible = False
                TocaAcordesEscalaMi.Visible = False
                TocaAcordesEscalaSi.Visible = False
                TocaAcordesEscalaFáSus.Visible = False
                TocaAcordesEscalaDóSus.Visible = False
                TocaAcordesEscalaSolSus.Visible = False
                TocaAcordesEscalaRéSus.Visible = False
                TocaAcordesEscalaLáSus.Visible = False
                TocaAcordesEscalaFá.Visible = False
            End If

            Me.SetBitmap(FaceBit, TransAmount)
            Teclado.Activate()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub Escalas_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        SalvaSettings()
    End Sub

    Public Sub SalvaSettings()

        Try

            My.Settings.NovaCor = ColorDialog1.Color
            My.Settings.NovaCorC = ColorDialogC.Color
            My.Settings.NovaCorD = ColorDialogD.Color
            My.Settings.NovaCorE = ColorDialogE.Color
            My.Settings.NovaCorF = ColorDialogF.Color
            My.Settings.NovaCorG = ColorDialogG.Color
            My.Settings.NovaCorA = ColorDialogA.Color
            My.Settings.NovaCorB = ColorDialogB.Color

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub ab_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown, _
      PictureBox6.MouseDown, CorC.MouseDown, CorD.MouseDown, CorE.MouseDown, CorF.MouseDown, CorG.MouseDown, CorA.MouseDown, CorB.MouseDown, _
     PictureBox2.MouseDown, PictureBox3.MouseDown, PictureBox4.MouseDown, PictureBox7.MouseDown
        VGa = Me.MousePosition.X - Me.Location.X
        VGb = Me.MousePosition.Y - Me.Location.Y
    End Sub

    Private Sub ab_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove, _
      PictureBox6.MouseMove, CorC.MouseMove, CorD.MouseMove, CorE.MouseMove, CorF.MouseMove, CorG.MouseMove, CorA.MouseMove, CorB.MouseMove, _
     PictureBox2.MouseMove, PictureBox3.MouseMove, PictureBox4.MouseMove, PictureBox7.MouseMove

        If e.Button = MouseButtons.Left Then
            VGnewPoint = Me.MousePosition
            VGnewPoint.X -= VGa
            VGnewPoint.Y -= VGb
            Me.Location = VGnewPoint
            Teclado.Top = Me.Top + 674
            Teclado.Left = Me.Left
        End If

    End Sub

    Private Sub Escalas_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseUp, _
     PictureBox6.MouseUp, CorC.MouseUp, CorD.MouseUp, CorE.MouseUp, CorF.MouseUp, CorG.MouseUp, CorA.MouseUp, CorB.MouseUp, _
     PictureBox2.MouseUp, PictureBox3.MouseUp, PictureBox4.MouseUp, PictureBox7.MouseUp
        Teclado.Activate()
    End Sub

    Private Sub Escalas_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try

            Me.Top -= 115
            NomeEscala = ""

            DecimaThread = New Thread(AddressOf DecimaThreadCode)
            DecimaThread.Name = "Decima Thread"
            DecimaThread.Start()


            ToolTip1.SetToolTip(PictureBox2, "Em azul a ordem em que aparecem os sustenidos e bemóis." & vbCrLf & vbCrLf & "Sustenidos --> F#, C#, G#, D#, A#, E#, B#" & vbCrLf & vbCrLf & "Bemóis --> Bb, Eb, Ab, Db, Gb, Cb, Fb")

            Me.BringToFront()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub DecimaThreadCode()
        AnulaVariaveisDosIntervalos()
        Teclado.DefineTecladoInicial()
    End Sub

    Public Sub AnulaVariaveisDosIntervalos()

        Try

            'anula variáveis dos intervalos
            I_1_0 = 100 : I_a1_1 = 100 : I_b2_2 = 100 : I_2_3 = 100 : I_a2_4 = 100 : I_b3_5 = 100 : I_3_6 = 100
            I_a3_7 = 100 : I_b4_8 = 100 : I_4_9 = 100 : I_a4_10 = 100 : I_b5_11 = 100 : I_5_12 = 100 : I_a5_13 = 100
            I_b6_14 = 100 : I_6_15 = 100 : I_a6_16 = 100 : I_b7_17 = 100 : I_7_18 = 100 : I_a7_19 = 100 : I_b8_20 = 100 : I_8_21 = 100

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Public Sub GeraEscalas()

        Try

            Teclado.atualizaTeclado()
            VGjjj = ""
            Teclado.ListBox1.Items.Clear()


            'zera as variáveis que somam os intervalos 
            VGc = 0 : VGd = 0 : VGee = 0 : VGf = 0 : VGg = 0 : VGh = 0 : VGi = 0 : VGj = 0 : VGk = 0 : VGl = 0 : VGm = 0 : VGn = 0 : VGt = 0

            AnulaVariaveisDosIntervalos()


            If VGoo = "Cromática (s-s-s-s-s-s-s-s-s-s-s-s-s)" Then
                VGc = 1 : VGd = 1 : VGee = 1 : VGf = 1 : VGg = 1 : VGh = 1 : VGi = 1 : VGj = 1 : VGk = 1 : VGl = 1 : VGm = 1 : VGn = 1 : qtdeloop = 13
                I_1_0 = 0 : I_b2_2 = 2 : I_2_3 = 3 : I_a2_4 = 4 : I_3_6 = 6 : I_4_9 = 9 : I_a4_10 = 10 : I_5_12 = 12 : I_a5_13 = 13 : I_6_15 = 15 : I_a6_16 = 16 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Maior (T-T-s-T-T-T-s)" OrElse VGoo = "Jônio (T-T-s-T-T-T-s)" OrElse VGoo = "Etíope 1 ou Ethiopian (A raray) (T-T-s-T-T-T-s)" OrElse VGoo = "Ethiopian (A raray)" _
                     OrElse VGoo = "Mela Dhirasankarabharana (29) (T-T-s-T-T-T-s)" OrElse VGoo = "Theta, Bilaval (T-T-s-T-T-T-s)" OrElse VGoo = "Ghana Heptatônica (T-T-s-T-T-T-s)" _
                     OrElse VGoo = "Peruana maior (T-T-s-T-T-T-s)" Then
                VGc = 2 : VGd = 2 : VGee = 1 : VGf = 2 : VGg = 2 : VGh = 2 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Maior", "Jônio", "Etíope 1 ou Ethiopian (A raray)", "Mela Dhirasankarabharana (29)", "Theta, Bilaval", "Ghana heptatônica", "Peruana maior"})
            ElseIf VGoo = "Menor Natural ou Menor Pura (T-s-T-T-s-T-T)" OrElse VGoo = "Eólio (T-s-T-T-s-T-T)" OrElse VGoo = "Menor Melódica - Descendente (T-s-T-T-s-T-T)" OrElse VGoo = "Etíope 2 ou Ethiopian (Geez && Ezel) (T-s-T-T-s-T-T)" OrElse _
           VGoo = "Mela Natabhairavi (20) (T-s-T-T-s-T-T)" OrElse VGoo = "Theta, Asavari (T-s-T-T-s-T-T)" OrElse VGoo = "Raga Adana (T-s-T-T-s-T-T)" OrElse VGoo = "Peruana menor (T-s-T-T-s-T-T)" OrElse _
           VGoo = "Chiao (T-s-T-T-s-T-T)" Then
                VGc = 2 : VGd = 1 : VGee = 2 : VGf = 2 : VGg = 1 : VGh = 2 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Menor Natural ou Menor Pura", "Eólio", "Menor Melódica Descendente", "Etíope 2 ou Ethiopian (Geez & Ezel)", "Mela Natabhairavi (20)", "Theta, Asavari", "Raga Adana", "Peruana menor", "Chiao"})
            ElseIf VGoo = "Menor Melódica - Ascendente (T-s-T-T-T-T-s)" OrElse VGoo = "Jazz Menor (T-s-T-T-T-T-s)" OrElse VGoo = "Hawayana 2 (T-s-T-T-T-T-s)" OrElse VGoo = "Mela Gaurimanohari (23) (T-s-T-T-T-T-s)" _
             OrElse VGoo = "Mischung 1 (T-s-T-T-T-T-s)" OrElse VGoo = "Raga Patdip (T-s-T-T-T-T-s)" Then
                VGc = 2 : VGd = 1 : VGee = 2 : VGf = 2 : VGg = 2 : VGh = 2 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Menor Melódica", "Jazz Menor", "Hawayana 2", "Mela Gaurimanohari (23)", "Mischung 1", "Raga Patdip"})
            ElseIf VGoo = "Menor Harmônica (T-s-T-T-s-T½-s)" OrElse VGoo = "Cíngara ou Zíngara Espanhola (T-s-T-T-s-T½-s)" OrElse VGoo = "Mela Kiravani (21) (T-s-T-T-s-T½-s)" OrElse VGoo = "Maometana (T-s-T-T-s-T½-s)" _
             OrElse VGoo = "Maqam Bayat-e-Esfahan (T-s-T-T-s-T½-s)" OrElse VGoo = "Mischung 4 (T-s-T-T-s-T½-s)" OrElse VGoo = "Raga Kiranavali (T-s-T-T-s-T½-s)" Then
                VGc = 2 : VGd = 1 : VGee = 2 : VGf = 2 : VGg = 1 : VGh = 3 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Menor Harmônica", "Cíngara Espanhola", "Mela Kiravani (21)", "Maometana", "Maqam Bayat-e-Esfahan", "Mischung 4", "Raga Kiranavali"})
            ElseIf VGoo = "Dominant 7th (T-T-s-T-T-s-T)" OrElse VGoo = "Mixolídio (T-T-s-T-T-s-T)" OrElse VGoo = "Mela Harikambhoji (28) (T-T-s-T-T-s-T)" OrElse VGoo = "Theta, Khamaj (T-T-s-T-T-s-T)" _
             OrElse VGoo = "Mischung 3 (T-T-s-T-T-s-T)" Then
                VGc = 2 : VGd = 2 : VGee = 1 : VGf = 2 : VGg = 2 : VGh = 1 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Mixolídio", "Dominant 7th", "Mela Harikambhoji (28)", "Theta, Khamaj", "Mischung 3"})
            ElseIf VGoo = "Pentatônica Maior 1 (T-T-T½-T-T½)" OrElse VGoo = "Mongol Chinesa (T-T-T½-T-T½)" OrElse VGoo = "Diatônica (T-T-T½-T-T½)" OrElse VGoo = "Coreana 1 (T-T-T½-T-T½)" _
            OrElse VGoo = "Ghana Pentatônica 2 (T-T-T½-T-T½)" OrElse VGoo = "Gong (T-T-T½-T-T½)" OrElse VGoo = "Ryosen (T-T-T½-T-T½)" OrElse VGoo = "China 2 (T-T-T½-T-T½)" Then
                VGc = 2 : VGd = 2 : VGee = 3 : VGf = 2 : VGg = 3 : qtdeloop = 6
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_5_12 = 12 : I_6_15 = 15 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Pentatônica Maior", "Mongol Chinesa", "Diatônica", "Coreana 1", "Ghana pentatônica 2", "Gong", "Ryosen", "China 2"})
            ElseIf VGoo = "Pentatônica Menor 1 (T½-T-T-T½-T)" OrElse VGoo = "Minyo (T½-T-T-T½-T)" OrElse VGoo = "Yu pentatônica (T½-T-T-T½-T)" Then
                VGc = 3 : VGd = 2 : VGee = 2 : VGf = 3 : VGg = 2 : qtdeloop = 6
                I_1_0 = 0 : I_b3_5 = 5 : I_4_9 = 9 : I_5_12 = 12 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Pentatônica Menor", "Minyo", "Yu pentatônica"})
            ElseIf VGoo = "Pentatônica Blues (T½-T-s-s-T½-T)" OrElse VGoo = "Blues Menor (T½-T-s-s-T½-T)" OrElse VGoo = "Blues 1 (T½-T-s-s-T½-T)" Then
                VGc = 3 : VGd = 2 : VGee = 1 : VGf = 1 : VGg = 3 : VGh = 2 : qtdeloop = 7
                I_1_0 = 0 : I_b3_5 = 5 : I_4_9 = 9 : I_b5_11 = 11 : I_5_12 = 12 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Pentatônica Blues", "Blues Menor", "Blues 1"})
            ElseIf VGoo = "Pentatônica Neutra 1 (T-T½-T-T½-T)" OrElse VGoo = "Egípcia (T-T½-T-T½-T)" OrElse VGoo = "Jin Yu (T-T½-T-T½-T)" OrElse VGoo = "Raga Madhyamavati (T-T½-T-T½-T)" OrElse VGoo = "Yo (T-T½-T-T½-T)" Then
                VGc = 2 : VGd = 3 : VGee = 2 : VGf = 3 : VGg = 2 : qtdeloop = 6
                I_1_0 = 0 : I_2_3 = 3 : I_4_9 = 9 : I_5_12 = 12 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Pentatônica Neutra", "Egípcia", "Jin Yu", "Raga Madhyamavati", "Yo"})
            ElseIf VGoo = "Dório (T-s-T-T-T-s-T)" OrElse VGoo = "Mela Kharaharapriya (22) (T-s-T-T-T-s-T)" OrElse VGoo = "Theta, Kafi (T-s-T-T-T-s-T)" OrElse VGoo = "Esquimal Heptatônica (T-s-T-T-T-s-T)" _
             OrElse VGoo = "Mischung 5 (T-s-T-T-T-s-T)" OrElse VGoo = "Yu heptatônica (T-s-T-T-T-s-T)" Then
                VGc = 2 : VGd = 1 : VGee = 2 : VGf = 2 : VGg = 2 : VGh = 1 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Dório", "Mela Kharaharapriya (22)", "Theta, Kafi", "Esquimal heptatônica", "Mischung 5", "Yu heptatônica"})
            ElseIf VGoo = "Frígio (s-T-T-T-s-T-T)" OrElse VGoo = "Mela Hanumattodi (8) (s-T-T-T-s-T-T)" OrElse VGoo = "Napolitana Menor (s-T-T-T-s-T-T)" OrElse VGoo = "Theta, Bhairavi (s-T-T-T-s-T-T)" OrElse VGoo = "In (s-T-T-T-s-T-T)" _
             OrElse VGoo = "Maior Invertida (s-T-T-T-s-T-T)" OrElse VGoo = "Maqam Kurd (s-T-T-T-s-T-T)" OrElse VGoo = "Ousak (s-T-T-T-s-T-T)" OrElse VGoo = "Raga Bilashkhani Todi (s-T-T-T-s-T-T)" Then
                VGc = 1 : VGd = 2 : VGee = 2 : VGf = 2 : VGg = 1 : VGh = 2 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Frígio", "Mela Hanumattodi (8)", "Napolitana Menor", "Theta, Bhairavi", "In", "Maior Invertida", "Maqam Kurd", "Ousak", "Raga Bilashkhani Todi"})
            ElseIf VGoo = "Lídio (T-T-T-s-T-T-s)" OrElse VGoo = "Mela Mechakalyani (65) (T-T-T-s-T-T-s)" OrElse VGoo = "Theta, Kalyan (T-T-T-s-T-T-s)" OrElse VGoo = "Raga Shuddh Kalyan (T-T-T-s-T-T-s)" Then
                VGc = 2 : VGd = 2 : VGee = 2 : VGf = 1 : VGg = 2 : VGh = 2 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_a4_10 = 10 : I_5_12 = 12 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Lídio", "Mela Mechakalyani (65)", "Theta, Kalyan", "Raga Shuddh Kalyan"})
            ElseIf VGoo = "Lócrio (s-T-T-s-T-T-T)" OrElse VGoo = "Pien Chih (s-T-T-s-T-T-T)" Then
                VGc = 1 : VGd = 2 : VGee = 2 : VGf = 1 : VGg = 2 : VGh = 2 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_4_9 = 9 : I_b5_11 = 11 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Lócrio", "Pien Chih"})
            ElseIf VGoo = "Tons Inteiros Diminuida (s-T-s-T-T-T-T)" OrElse VGoo = "Super Lócrio (s-T-s-T-T-T-T)" Then
                VGc = 1 : VGd = 2 : VGee = 1 : VGf = 2 : VGg = 2 : VGh = 2 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_a2_4 = 4 : I_3_6 = 6 : I_b5_11 = 11 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Tons Inteiros Diminuida", "Super Lócrio"})
            ElseIf VGoo = "Tons Inteiros (T-T-T-T-T-T)" OrElse VGoo = "Auxiliar Aumentada (T-T-T-T-T-T)" OrElse VGoo = "Aumentada 3 (T-T-T-T-T-T)" OrElse VGoo = "Raga Gopriya (T-T-T-T-T-T)" OrElse VGoo = "Tonal (T-T-T-T-T-T)" Then
                VGc = 2 : VGd = 2 : VGee = 2 : VGf = 2 : VGg = 2 : VGh = 2 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_a4_10 = 10 : I_a5_13 = 13 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Tons Inteiros", "Auxiliar Aumentada", "Aumentada 3", "Raga Gopriya", "Tonal"})
            ElseIf VGoo = "Leading Whole Tone (T-T-T-T-T-s-s)" Then
                VGc = 2 : VGd = 2 : VGee = 2 : VGf = 2 : VGg = 2 : VGh = 1 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_a4_10 = 10 : I_a5_13 = 13 : I_a6_16 = 16 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Cigana Espanhola (s-T½-s-T-s-T-T)" OrElse VGoo = "Judaica (Ahaba Rabba) (s-T½-s-T-s-T-T)" OrElse VGoo = "Ahavoh Rabboh (s-T½-s-T-s-T-T)" OrElse VGoo = "Alhijaz (s-T½-s-T-s-T-T)" _
             OrElse VGoo = "Bizantina 1 (s-T½-s-T-s-T-T)" OrElse VGoo = "Flamenca (s-T½-s-T-s-T-T)" OrElse VGoo = "Frigio Maior (s-T½-s-T-s-T-T)" OrElse VGoo = "Hitzaz (s-T½-s-T-s-T-T)" OrElse VGoo = "Israelita 2 (s-T½-s-T-s-T-T)" _
             OrElse VGoo = "Maqam Humayun (s-T½-s-T-s-T-T)" Then
                VGc = 1 : VGd = 3 : VGee = 1 : VGf = 2 : VGg = 1 : VGh = 2 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Cigana Espanhola", "Judaica (Ahaba Rabba)", "Ahavoh Rabboh", "Alhijaz", "Bizantina 1", "Flamenca", "Frigio Maior", "Hitzaz", "Israelita 2", "Maqam Humayun"})
            ElseIf VGoo = "Dupla Harmônica (s-T½-s-T-s-T½-s)" OrElse VGoo = "Bizantina 2 (s-T½-s-T-s-T½-s)" OrElse VGoo = "Hispano-árabe (s-T½-s-T-s-T½-s)" OrElse VGoo = "Hitzaskiar (s-T½-s-T-s-T½-s)" _
             OrElse VGoo = "Maqam Zengule (s-T½-s-T-s-T½-s)" OrElse VGoo = "Persa 2 (s-T½-s-T-s-T½-s)" OrElse VGoo = "Cíngara Maior 1 (s-T½-s-T-s-T½-s)" OrElse VGoo = "Theta, Bhairav (s-T½-s-T-s-T½-s)" OrElse VGoo = "Cigana Hungaro-Persa (s-T½-s-T-s-T½-s)" Then
                VGc = 1 : VGd = 3 : VGee = 1 : VGf = 2 : VGg = 1 : VGh = 3 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Dupla Harmônica", "Bizantina 2", "Hispano-árabe", "Hitzaskiar", "Maqam Zengule", "Persa 2", "Cíngara Maior 1", "Theta, Bhairav", "Cigana Hungaro-Persa"})
            ElseIf VGoo = "Octatônica (s-T) (s-T-s-T-s-T-s-T)" OrElse VGoo = "Auxiliar Diminuída Blues (s-T-s-T-s-T-s-T)" OrElse VGoo = "Diminuída, Half (s-T-s-T-s-T-s-T)" OrElse VGoo = "Blues 4 (s-T-s-T-s-T-s-T)" _
             OrElse VGoo = "Diminuída 3 (s-T-s-T-s-T-s-T)" OrElse VGoo = "Semitom-Tom (s-T-s-T-s-T-s-T)" OrElse VGoo = "Simétrica 2 (s-T-s-T-s-T-s-T)" Then
                VGc = 1 : VGd = 2 : VGee = 1 : VGf = 2 : VGg = 1 : VGh = 2 : VGi = 1 : VGj = 2 : qtdeloop = 9
                If VGoo = "Blues 4 (s-T-s-T-s-T-s-T)" Then I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_3_6 = 6 : I_b5_11 = 11 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
                If VGoo = "Diminuída 3 (s-T-s-T-s-T-s-T)" OrElse VGoo = "Octatônica (s-T) (s-T-s-T-s-T-s-T)" OrElse VGoo = "Semitom-Tom (s-T-s-T-s-T-s-T)" OrElse VGoo = "Auxiliar Diminuída Blues (s-T-s-T-s-T-s-T)" OrElse VGoo = "Diminuída, Half (s-T-s-T-s-T-s-T)" Then I_1_0 = 0 : I_b2_2 = 2 : I_a2_4 = 4 : I_3_6 = 6 : I_a4_10 = 10 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21 : Intervalos = "Intervalos: 1,b2,#2,3,#4,5,6,b7,8    (U,2m,2a,3M,4a,5J,6M,7m,8J)"
                If VGoo = "Simétrica 2 (s-T-s-T-s-T-s-T)" Then I_1_0 = 0 : I_a1_1 = 1 : I_a2_4 = 4 : I_3_6 = 6 : I_a4_10 = 10 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Octatônicas (s-T)", "Auxiliar Diminuída Blues", "Diminuída, Half", "Blues 4", "Diminuída 3", "Semitom-Tom", "Simétrica 2"})
            ElseIf VGoo = "Raga Ramkali (s-T½-s-s-s-s-T½-s)" Then
                VGc = 1 : VGd = 3 : VGee = 1 : VGf = 1 : VGg = 1 : VGh = 1 : VGi = 3 : VGj = 1 : qtdeloop = 9
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_4_9 = 9 : I_b5_11 = 11 : I_5_12 = 12 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Blues Maior (T-s-s-T½-T-T½)" Then
                VGc = 2 : VGd = 1 : VGee = 1 : VGf = 3 : VGg = 2 : VGh = 3 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_3_6 = 6 : I_5_12 = 12 : I_6_15 = 15 : I_8_21 = 21
            ElseIf VGoo = "Blues 3 (T-s-s-s-s-s-T-s-s-s)" Then
                VGc = 2 : VGd = 1 : VGee = 1 : VGf = 1 : VGg = 1 : VGh = 1 : VGi = 2 : VGj = 1 : VGk = 1 : VGl = 1 : qtdeloop = 11
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_3_6 = 6 : I_4_9 = 9 : I_a4_10 = 10 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Blues 8 (T-s-s-s-s-s-T-s-T)" Then
                VGc = 2 : VGd = 1 : VGee = 1 : VGf = 1 : VGg = 1 : VGh = 1 : VGi = 2 : VGj = 1 : VGk = 2 : qtdeloop = 10
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_3_6 = 6 : I_4_9 = 9 : I_b5_11 = 11 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Blues 6 (T-s-T-s-s-T-s-T)" Then
                VGc = 2 : VGd = 1 : VGee = 2 : VGf = 1 : VGg = 1 : VGh = 2 : VGi = 1 : VGj = 2 : qtdeloop = 9
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_4_9 = 9 : I_b5_11 = 11 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Blues 5 (T-s-T-s-s-T½-T)" Then
                VGc = 2 : VGd = 1 : VGee = 2 : VGf = 1 : VGg = 1 : VGh = 3 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_4_9 = 9 : I_b5_11 = 11 : I_5_12 = 12 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Blues 9 (T½-s-s-s-s-T-s-s-s)" Then
                VGc = 3 : VGd = 1 : VGee = 1 : VGf = 1 : VGg = 1 : VGh = 2 : VGi = 1 : VGj = 1 : VGk = 1 : qtdeloop = 10
                I_1_0 = 0 : I_b3_5 = 5 : I_3_6 = 6 : I_4_9 = 9 : I_b5_11 = 11 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Blues 7 (T½-s-s-s-s-T½-s-s)" Then
                VGc = 3 : VGd = 1 : VGee = 1 : VGf = 1 : VGg = 1 : VGh = 3 : VGi = 1 : VGj = 1 : qtdeloop = 9
                I_1_0 = 0 : I_b3_5 = 5 : I_3_6 = 6 : I_4_9 = 9 : I_b5_11 = 11 : I_5_12 = 12 : I_b7_17 = 17 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Mela Vakulabharanam (14) (s-T½-s-T-s-T-T)" OrElse VGoo = "Mela Ragavardhani (32) (T½-s-s-T-s-T-T)" OrElse VGoo = "Blues 10 (T½-s-s-T-s-T-T)" Then
                VGc = 3 : VGd = 1 : VGee = 1 : VGf = 2 : VGg = 1 : VGh = 2 : VGi = 2 : qtdeloop = 8
                If VGoo = "Blues 10 (T½-s-s-T-s-T-T)" Then I_1_0 = 0 : I_b3_5 = 5 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
                If VGoo = "Mela Ragavardhani (32) (T½-s-s-T-s-T-T)" OrElse VGoo = "Mela Vakulabharanam (14) (s-T½-s-T-s-T-T)" Then I_1_0 = 0 : I_a2_4 = 4 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Mela Vakulabharanam (14)", "Mela Ragavardhani (32)", "Blues 10"})
            ElseIf VGoo = "Blues 2 (T½-T-s-s-T½-s-s)" OrElse VGoo = "Hipofrigia cromática (T½-T-s-s-T½-s-s)" Then
                VGc = 3 : VGd = 2 : VGee = 1 : VGf = 1 : VGg = 3 : VGh = 1 : VGi = 1 : qtdeloop = 8
                If VGoo = "Blues 2 (T½-T-s-s-T½-s-s)" Then I_1_0 = 0 : I_b3_5 = 5 : I_4_9 = 9 : I_b5_11 = 11 : I_5_12 = 12 : I_b7_17 = 17 : I_7_18 = 18 : I_8_21 = 21
                If VGoo = "Hipofrigia cromática (T½-T-s-s-T½-s-s)" Then I_1_0 = 0 : I_b3_5 = 5 : I_4_9 = 9 : I_b5_11 = 11 : I_5_12 = 12 : I_a6_16 = 16 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Blues 2", "Hipofrigia cromática"})
            ElseIf VGoo = "Raga Suddha Bangala (T-s-T-T-T-T½)" Then
                VGc = 2 : VGd = 1 : VGee = 2 : VGf = 2 : VGg = 2 : VGh = 3 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_8_21 = 21
            ElseIf VGoo = "Raga Palasi (T-s-T-T-T½-T)" OrElse VGoo = "Menor Hexatônica (T-s-T-T-T½-T)" OrElse VGoo = "Esquimal Hexatônica 2 (T-s-T-T-T½-T)" Then
                VGc = 2 : VGd = 1 : VGee = 2 : VGf = 2 : VGg = 3 : VGh = 2 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_4_9 = 9 : I_5_12 = 12 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Raga Palasi", "Menor Hexatônica", "Esquimal Hexatônica 2"})
            ElseIf VGoo = "Raga Sindhura Kafi (T-s-T-T-2T-s)" Then
                VGc = 2 : VGd = 1 : VGee = 2 : VGf = 2 : VGg = 4 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_4_9 = 9 : I_5_12 = 12 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Purnalalita (T-s-T-T-2T½)" OrElse VGoo = "Ghana Pentatônica 1 (T-s-T-T-2T½)" OrElse VGoo = "Chad Gadyo (T-s-T-T-2T½)" Then
                VGc = 2 : VGd = 1 : VGee = 2 : VGf = 2 : VGg = 5 : qtdeloop = 6
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_4_9 = 9 : I_5_12 = 12 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Raga Purnalalita", "Ghana Pentatônica 1", "Chad Gadyo"})
            ElseIf VGoo = "Raga Ghantana (T-s-T-T½-T½-s)" Then
                VGc = 2 : VGd = 1 : VGee = 2 : VGf = 3 : VGg = 3 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_4_9 = 9 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Audav Tukhari (T-s-T-T½-2T)" Then
                VGc = 2 : VGd = 1 : VGee = 2 : VGf = 3 : VGg = 4 : qtdeloop = 6
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_4_9 = 9 : I_a5_13 = 13 : I_8_21 = 21
            ElseIf VGoo = "Raga Bagesri (T-s-T-2T-s-T)" Then
                VGc = 2 : VGd = 1 : VGee = 2 : VGf = 4 : VGg = 1 : VGh = 2 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_4_9 = 9 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Raga Abhogi (T-s-T-2T-T½)" Then
                VGc = 2 : VGd = 1 : VGee = 2 : VGf = 4 : VGg = 3 : qtdeloop = 6
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_4_9 = 9 : I_6_15 = 15 : I_8_21 = 21
            ElseIf VGoo = "Raga Cintamani (T-s-T½-s-s-s-s-T)" Then
                VGc = 2 : VGd = 1 : VGee = 3 : VGf = 1 : VGg = 1 : VGh = 1 : VGi = 1 : VGj = 2 : qtdeloop = 9
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_a4_10 = 10 : I_5_12 = 12 : I_b6_14 = 14 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Raga Syamalam (T-s-T½-s-s-2T)" Then
                VGc = 2 : VGd = 1 : VGee = 3 : VGf = 1 : VGg = 1 : VGh = 4 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_a4_10 = 10 : I_5_12 = 12 : I_b6_14 = 14 : I_8_21 = 21
            ElseIf VGoo = "Raga Vijayanagari (T-s-T½-s-T-T½)" Then
                VGc = 2 : VGd = 1 : VGee = 3 : VGf = 1 : VGg = 2 : VGh = 3 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_a4_10 = 10 : I_5_12 = 12 : I_6_15 = 15 : I_8_21 = 21
            ElseIf VGoo = "Raga Simharava (T-s-T½-s-T½-T)" Then
                VGc = 2 : VGd = 1 : VGee = 3 : VGf = 1 : VGg = 3 : VGh = 2 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_a4_10 = 10 : I_5_12 = 12 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Raga Amarasenapriya (T-s-T½-s-2T-s)" OrElse VGoo = "Raga Kaikavasi (T-s-T½-s-2T-s)" Then
                VGc = 2 : VGd = 1 : VGee = 3 : VGf = 1 : VGg = 4 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_a4_10 = 10 : I_5_12 = 12 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Raga Amarasenapriya", "Raga Kaikavasi"})
            ElseIf VGoo = "Raga Ranjani (T-s-T½-T½-T-s)" Then
                VGc = 2 : VGd = 1 : VGee = 3 : VGf = 3 : VGg = 2 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_b5_11 = 11 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Trimurti (T-s-2T-s-T-T)" Then
                VGc = 2 : VGd = 1 : VGee = 4 : VGf = 1 : VGg = 2 : VGh = 2 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_5_12 = 12 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Raga Manavi (T-s-2T-T-s-T)" Then
                VGc = 2 : VGd = 1 : VGee = 4 : VGf = 2 : VGg = 1 : VGh = 2 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Raga Sivaranjini (T-s-2T-T-T½)" OrElse VGoo = "Pentatônica Menor 2 (T-s-2T-T-T½)" OrElse VGoo = "Akebono (T-s-2T-T-T½)" OrElse VGoo = "Kumoi (T-s-2T-T-T½)" Then
                VGc = 2 : VGd = 1 : VGee = 4 : VGf = 2 : VGg = 3 : qtdeloop = 6
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_5_12 = 12 : I_6_15 = 15 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Raga Sivaranjini", "Pentatônica Menor 2", "Akebono", "Kumoi Scale"})
            ElseIf VGoo = "Mela Bhavapriya (44) (s-s-T½-T-s-s-T½)" OrElse VGoo = "Mela Kanakangi (1) (s-s-T½-T-s-s-T½)" OrElse VGoo = "Dórica cromática (s-s-T½-T-s-s-T½)" OrElse VGoo = "Raga Kanakambari (s-s-T½-T-s-s-T½)" Then
                VGc = 1 : VGd = 1 : VGee = 3 : VGf = 2 : VGg = 1 : VGh = 1 : VGi = 3 : qtdeloop = 8
                If VGoo = "Raga Kanakambari (s-s-T½-T-s-s-T½)" Then I_1_0 = 0 : I_b2_2 = 2 : I_2_3 = 3 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_6_15 = 15 : I_8_21 = 21
                If VGoo = "Dórica cromática (s-s-T½-T-s-s-T½)" Then I_1_0 = 0 : I_b2_2 = 2 : I_2_3 = 3 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_6_15 = 15 : I_8_21 = 21
                If VGoo = "Mela Bhavapriya (44) (s-s-T½-T-s-s-T½)" OrElse VGoo = "Mela Kanakangi (1) (s-s-T½-T-s-s-T½)" Then I_1_0 = 0 : I_b2_2 = 2 : I_2_3 = 3 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_6_15 = 15 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Mela Bhavapriya (44)", "Mela Kanakangi (1)", "Dórica cromática", "Raga Kanakambari"})
            ElseIf VGoo = "Javanesa 1 (s-T-T-T-T-s-T)" OrElse VGoo = "Mela Natakapriya (10) (s-T-T-T-T-s-T)" OrElse VGoo = "Dórica b2 (s-T-T-T-T-s-T)" OrElse VGoo = "Frigia #6 (s-T-T-T-T-s-T)" Then
                VGc = 1 : VGd = 2 : VGee = 2 : VGf = 2 : VGg = 2 : VGh = 1 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Javanesa", "Mela Natakapriya (10)", "Dórica b2", "Frigia #6"})
            ElseIf VGoo = "Javanesa 3 (s-T-2T-s-2T)" OrElse VGoo = "Balinesa 2 (s-T-2T-s-2T)" OrElse VGoo = "Pelog (s-T-2T-s-2T)" OrElse VGoo = "Raga Bhupalam (s-T-2T-s-2T)" Then
                VGc = 1 : VGd = 2 : VGee = 4 : VGf = 1 : VGg = 4 : qtdeloop = 6
                I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_5_12 = 12 : I_b6_14 = 14 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Javanesa 3", "Balinesa 2", "Pelog", "Raga Bhupalam"})

            ElseIf VGoo = "Raga Sindhi Bhairavi (s-s-s-s-s-T-s-T-s-s)" Then
                VGc = 1 : VGd = 1 : VGee = 1 : VGf = 1 : VGg = 1 : VGh = 2 : VGi = 1 : VGj = 2 : VGk = 1 : VGl = 1 : qtdeloop = 11
                I_1_0 = 0 : I_b2_2 = 2 : I_2_3 = 3 : I_b3_5 = 5 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_b7_17 = 17 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Messiânica 5 (s-s-s-s-T-s-s-s-s-T)" Then
                VGc = 1 : VGd = 1 : VGee = 1 : VGf = 1 : VGg = 2 : VGh = 1 : VGi = 1 : VGj = 1 : VGk = 1 : VGl = 2 : qtdeloop = 11
                I_1_0 = 0 : I_b2_2 = 2 : I_2_3 = 3 : I_b3_5 = 5 : I_3_6 = 6 : I_a4_10 = 10 : I_5_12 = 12 : I_a5_13 = 13 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Judaica (Adonai Malakh) (s-s-s-T-T-T-s-T)" Then
                VGc = 1 : VGd = 1 : VGee = 1 : VGf = 2 : VGg = 2 : VGh = 2 : VGi = 1 : VGj = 2 : qtdeloop = 9
                I_1_0 = 0 : I_b2_2 = 2 : I_2_3 = 3 : I_b3_5 = 5 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Messiânica 2 (s-s-s-T½-s-s-s-T½)" Then
                VGc = 1 : VGd = 1 : VGee = 1 : VGf = 3 : VGg = 1 : VGh = 1 : VGi = 1 : VGj = 3 : qtdeloop = 9
                I_1_0 = 0 : I_b2_2 = 2 : I_2_3 = 3 : I_b3_5 = 5 : I_a4_10 = 10 : I_5_12 = 12 : I_a5_13 = 13 : I_6_15 = 15 : I_8_21 = 21
            ElseIf VGoo = "Simétrica 3 (s-s-T-s-s-s-s-T-s-s)" Then
                VGc = 1 : VGd = 1 : VGee = 2 : VGf = 1 : VGg = 1 : VGh = 1 : VGi = 1 : VGj = 2 : VGk = 1 : VGl = 1 : qtdeloop = 11
                I_1_0 = 0 : I_a1_1 = 1 : I_2_3 = 3 : I_3_6 = 6 : I_4_9 = 9 : I_a4_10 = 10 : I_5_12 = 12 : I_b6_14 = 14 : I_b7_17 = 17 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Messiânica 1 (s-s-T-s-s-s-T-s-T)" OrElse VGoo = "Youlan (s-s-T-s-s-s-T-s-T)" Then
                VGc = 1 : VGd = 1 : VGee = 2 : VGf = 1 : VGg = 1 : VGh = 1 : VGi = 2 : VGj = 1 : VGk = 2 : qtdeloop = 10
                I_1_0 = 0 : I_b2_2 = 2 : I_2_3 = 3 : I_3_6 = 6 : I_4_9 = 9 : I_b5_11 = 11 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Messiânica 1", "Youlan"})
            ElseIf VGoo = "Messiânica 4 (s-s-T-T-s-s-T-T)" Then
                VGc = 1 : VGd = 1 : VGee = 2 : VGf = 2 : VGg = 1 : VGh = 1 : VGi = 2 : VGj = 2 : qtdeloop = 9
                I_1_0 = 0 : I_b2_2 = 2 : I_2_3 = 3 : I_3_6 = 6 : I_a4_10 = 10 : I_5_12 = 12 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Mixolidia cromática (s-s-T½-s-s-T½-T)" Then
                VGc = 1 : VGd = 1 : VGee = 3 : VGf = 1 : VGg = 1 : VGh = 3 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_2_3 = 3 : I_4_9 = 9 : I_b5_11 = 11 : I_5_12 = 12 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Mela Ratnangi (2) (s-s-T½-T-s-T-T)" Then
                VGc = 1 : VGd = 1 : VGee = 3 : VGf = 2 : VGg = 1 : VGh = 2 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_2_3 = 3 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Mela Ganamurti (3) (s-s-T½-T-s-T½-s)" OrElse VGoo = "Raga Ganasamavarali (s-s-T½-T-s-T½-s)" Then
                VGc = 1 : VGd = 1 : VGee = 3 : VGf = 2 : VGg = 1 : VGh = 3 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_2_3 = 3 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Mela Ganamurti (3)", "Raga Ganasamavarali"})
            ElseIf VGoo = "Mela Vanaspati (4) (s-s-T½-T-T-s-T)" OrElse VGoo = "Raga Bhanumati (s-s-T½-T-T-s-T)" Then
                VGc = 1 : VGd = 1 : VGee = 3 : VGf = 2 : VGg = 2 : VGh = 1 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_2_3 = 3 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Mela Vanaspati (4)", "Raga Bhanumati"})
            ElseIf VGoo = "Mela Manavati (5) (s-s-T½-T-T-T-s)" Then
                VGc = 1 : VGd = 1 : VGee = 3 : VGf = 2 : VGg = 2 : VGh = 2 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_2_3 = 3 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Mela Tanarupi (6) (s-s-T½-T-T½-s-s)" Then
                VGc = 1 : VGd = 1 : VGee = 3 : VGf = 2 : VGg = 3 : VGh = 1 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_2_3 = 3 : I_4_9 = 9 : I_5_12 = 12 : I_a6_16 = 16 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Suddha Mukhari (s-s-T½-T½-s-T½)" Then
                VGc = 1 : VGd = 1 : VGee = 3 : VGf = 3 : VGg = 1 : VGh = 3 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_2_3 = 3 : I_4_9 = 9 : I_b6_14 = 14 : I_6_15 = 15 : I_8_21 = 21
            ElseIf VGoo = "Mela Salagam (37) (s-s-2T-s-s-s-T½)" Then
                VGc = 1 : VGd = 1 : VGee = 4 : VGf = 1 : VGg = 1 : VGh = 2 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_2_3 = 3 : I_a4_10 = 10 : I_5_12 = 12 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Mela Jhalavarali (39) (s-s-2T-s-s-T½-s)" Then
                VGc = 1 : VGd = 1 : VGee = 4 : VGf = 1 : VGg = 1 : VGh = 3 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_2_3 = 3 : I_a4_10 = 10 : I_5_12 = 12 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Messiânica 3 (s-s-2T-s-s-2T)" Then
                VGc = 1 : VGd = 1 : VGee = 4 : VGf = 1 : VGg = 1 : VGh = 4 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_2_3 = 3 : I_a4_10 = 10 : I_5_12 = 12 : I_b6_14 = 14 : I_8_21 = 21
            ElseIf VGoo = "Mela Navanitam (40) (s-s-2T-s-T-s-T)" Then
                VGc = 1 : VGd = 1 : VGee = 4 : VGf = 1 : VGg = 2 : VGh = 1 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_2_3 = 3 : I_a4_10 = 10 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Mela Pavani (41) (s-s-2T-s-T-T-s)" OrElse VGoo = "Mela Suvarnangi (47) (s-s-2T-s-T-T-s)" Then
                VGc = 1 : VGd = 1 : VGee = 4 : VGf = 1 : VGg = 2 : VGh = 2 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_2_3 = 3 : I_a4_10 = 10 : I_5_12 = 12 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Mela Pavani (41)", "Mela Suvarnangi (47)"})
            ElseIf VGoo = "Raga Chandrajyoti (s-s-2T-s-T-T½)" Then
                VGc = 1 : VGd = 1 : VGee = 4 : VGf = 1 : VGg = 2 : VGh = 3 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_2_3 = 3 : I_a4_10 = 10 : I_5_12 = 12 : I_6_15 = 15 : I_8_21 = 21
            ElseIf VGoo = "Mela Raghupriya (42) (s-s-2T-s-T½-s-s)" Then
                VGc = 1 : VGd = 1 : VGee = 4 : VGf = 1 : VGg = 3 : VGh = 1 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_2_3 = 3 : I_a4_10 = 10 : I_5_12 = 12 : I_a6_16 = 16 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Nabhomani (s-s-2T-s-2T½)" Then
                VGc = 1 : VGd = 1 : VGee = 4 : VGf = 1 : VGg = 5 : qtdeloop = 6
                I_1_0 = 0 : I_b2_2 = 2 : I_2_3 = 3 : I_a4_10 = 10 : I_5_12 = 12 : I_8_21 = 21
            ElseIf VGoo = "Warao ditônica (s-T)" Then
                VGc = 1 : VGd = 2 : qtdeloop = 3
                I_1_0 = 0 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Espanhola octatônica (s-T-s-s-s-T-T-T)" OrElse VGoo = "Espla (s-T-s-s-s-T-T-T)" Then
                VGc = 1 : VGd = 2 : VGee = 1 : VGf = 1 : VGg = 1 : VGh = 2 : VGi = 2 : VGj = 2 : qtdeloop = 9
                If VGoo = "Espanhola octatônica (s-T-s-s-s-T-T-T)" Then I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_3_6 = 6 : I_4_9 = 9 : I_b5_11 = 11 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
                If VGoo = "Espla (s-T-s-s-s-T-T-T)" Then I_1_0 = 0 : I_b2_2 = 2 : I_a2_4 = 4 : I_3_6 = 6 : I_4_9 = 9 : I_b5_11 = 11 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Espanhola octatônica", "Espla"})
            ElseIf VGoo = "Maqam Shadd'araban (s-T-s-s-s-T½-s-T)" Then
                VGc = 1 : VGd = 2 : VGee = 1 : VGf = 1 : VGg = 1 : VGh = 3 : VGi = 1 : VGj = 2 : qtdeloop = 9
                I_1_0 = 0 : I_b2_2 = 2 : I_a2_4 = 4 : I_3_6 = 6 : I_4_9 = 9 : I_b5_11 = 11 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Genus chromaticum (s-T-s-s-T-s-s-T-s)" OrElse VGoo = "Tcherepnin (s-T-s-s-T-s-s-T-s)" Then
                VGc = 1 : VGd = 2 : VGee = 1 : VGf = 1 : VGg = 2 : VGh = 1 : VGi = 1 : VGj = 2 : VGk = 1 : qtdeloop = 10
                If VGoo = "Genus chromaticum (s-T-s-s-T-s-s-T-s)" Then I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
                If VGoo = "Tcherepnin (s-T-s-s-T-s-s-T-s)" Then I_1_0 = 0 : I_a1_1 = 1 : I_a2_4 = 4 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_a5_13 = 13 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Genus chromaticum", "Tcherepnin"})
            ElseIf VGoo = "Árabe 5 (s-T-s-s-T-s-T-s-s)" OrElse VGoo = "Frigia árabe (s-T-s-s-T-s-T-s-s)" Then
                VGc = 1 : VGd = 2 : VGee = 1 : VGf = 1 : VGg = 2 : VGh = 1 : VGi = 2 : VGj = 1 : VGk = 1 : qtdeloop = 10
                I_1_0 = 0 : I_b2_2 = 2 : I_a2_4 = 4 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_b7_17 = 17 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Árabe 5", "Frígia árabe"})
            ElseIf VGoo = "Frigia espanhola (s-T-s-s-T-s-T-T)" Then
                VGc = 1 : VGd = 2 : VGee = 1 : VGf = 1 : VGg = 2 : VGh = 1 : VGi = 2 : VGj = 2 : qtdeloop = 9
                I_1_0 = 0 : I_b2_2 = 2 : I_a2_4 = 4 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Israelita 1 (s-T-s-T-T-s-T-s)" OrElse VGoo = "Magen Abot (s-T-s-T-T-s-T-s)" Then
                VGc = 1 : VGd = 2 : VGee = 1 : VGf = 2 : VGg = 2 : VGh = 1 : VGi = 2 : VGj = 1 : qtdeloop = 9
                If VGoo = "Israelita 1 (s-T-s-T-T-s-T-s)" Then I_1_0 = 0 : I_a1_1 = 1 : I_a2_4 = 4 : I_3_6 = 6 : I_a4_10 = 10 : I_a5_13 = 13 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
                If VGoo = "Magen Abot (s-T-s-T-T-s-T-s)" Then I_1_0 = 0 : I_b2_2 = 2 : I_a2_4 = 4 : I_3_6 = 6 : I_a4_10 = 10 : I_a5_13 = 13 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Israelita 1", "Magen Abot"})
            ElseIf VGoo = "Ultra Lócria (s-T-s-T-T-s-T½)" Then
                VGc = 1 : VGd = 2 : VGee = 1 : VGf = 2 : VGg = 2 : VGh = 1 : VGi = 3 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_3_6 = 6 : I_b5_11 = 11 : I_b6_14 = 14 : I_6_15 = 15 : I_8_21 = 21
            ElseIf VGoo = "Maqam Huzzam (s-T-s-T½-s-T-T)" OrElse VGoo = "Frigia b4 (s-T-s-T½-s-T-T)" Then
                VGc = 1 : VGd = 2 : VGee = 1 : VGf = 3 : VGg = 1 : VGh = 2 : VGi = 2 : qtdeloop = 8
                If VGoo = "Maqam Huzzam (s-T-s-T½-s-T-T)" Then I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_3_6 = 6 : I_5_12 = 12 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
                If VGoo = "Frigia b4 (s-T-s-T½-s-T-T)" Then I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_b4_8 = 8 : I_5_12 = 12 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Maqam Huzzam", "Frígia b4"})
            ElseIf VGoo = "Be-Bop Semi-diminuída (s-T-T-s-s-s-T½-s)" Then
                VGc = 1 : VGd = 2 : VGee = 2 : VGf = 1 : VGg = 1 : VGh = 1 : VGi = 3 : VGj = 1 : qtdeloop = 9
                I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_4_9 = 9 : I_b5_11 = 11 : I_5_12 = 12 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Frigia doble hexatônica (s-T-T-s-T½-T½)" Then
                VGc = 1 : VGd = 2 : VGee = 2 : VGf = 1 : VGg = 3 : VGh = 3 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_4_9 = 9 : I_b5_11 = 11 : I_6_15 = 15 : I_8_21 = 21
            ElseIf VGoo = "Honchoshi Plagal (s-T-T-s-2T-T)" Then
                VGc = 1 : VGd = 2 : VGee = 2 : VGf = 1 : VGg = 4 : VGh = 2 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_4_9 = 9 : I_b5_11 = 11 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Mela Senavati (7) (s-T-T-T-s-s-T½)" Then
                VGc = 1 : VGd = 2 : VGee = 2 : VGf = 2 : VGg = 1 : VGh = 1 : VGi = 3 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_6_15 = 15 : I_8_21 = 21
            ElseIf VGoo = "Maqam Shahnaz Kurdi (s-T-T-T-s-T½-s)" OrElse VGoo = "Mela Dhenuka (9) (s-T-T-T-s-T½-s)" OrElse VGoo = "Napolitana 3 (s-T-T-T-s-T½-s)" Then
                VGc = 1 : VGd = 2 : VGee = 2 : VGf = 2 : VGg = 1 : VGh = 3 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Maqam Shahnaz Kurdi", "Mela Dhenuka (9)", "Napolitana 3"})
            ElseIf VGoo = "Raga Suddha Simantini (s-T-T-T-s-2T)" Then
                VGc = 1 : VGd = 2 : VGee = 2 : VGf = 2 : VGg = 1 : VGh = 4 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_8_21 = 21
            ElseIf VGoo = "Napolitana 2 (s-T-T-T-T-T-s)" OrElse VGoo = "Napolitana Maior (s-T-T-T-T-T-s)" OrElse VGoo = "Mela Kokilapriya (11) (s-T-T-T-T-T-s)" Then
                VGc = 1 : VGd = 2 : VGee = 2 : VGf = 2 : VGg = 2 : VGh = 2 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Napolitana 2", "Napolitana Maior", "Mela Kokilapriya (11)"})
            ElseIf VGoo = "Mela Rupavati (12) (s-T-T-T-T½-s-s)" Then
                VGc = 1 : VGd = 2 : VGee = 2 : VGf = 2 : VGg = 3 : VGh = 1 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_4_9 = 9 : I_5_12 = 12 : I_a6_16 = 16 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Gandharavam (s-T-T-T-T½-T)" Then
                VGc = 1 : VGd = 2 : VGee = 2 : VGf = 2 : VGg = 3 : VGh = 2 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_4_9 = 9 : I_5_12 = 12 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Raga Suddha Todi (s-T-T-T½-T-T)" OrElse VGoo = "Ritsu (s-T-T-T½-T-T)" Then
                VGc = 1 : VGd = 2 : VGee = 2 : VGf = 3 : VGg = 2 : VGh = 2 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_4_9 = 9 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Raga Suddha Todi", "Ritsu"})
            ElseIf VGoo = "Raga Viyogavarali (s-T-T-T½-T½-s)" Then
                VGc = 1 : VGd = 2 : VGee = 2 : VGf = 3 : VGg = 3 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_4_9 = 9 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Balinesa 1 (s-T-T-T½-2T)" OrElse VGoo = "Enigmática 1 (s-T-T-T½-2T)" OrElse VGoo = "Javanesa 2 (s-T-T-T½-2T)" Then
                VGc = 1 : VGd = 2 : VGee = 2 : VGf = 3 : VGg = 4 : qtdeloop = 6
                I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_4_9 = 9 : I_b6_14 = 14 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Balinesa 1", "Enigmática 1", "Javanesa 1"})
            ElseIf VGoo = "Mela Gavambodhi (43) (s-T-T½-s-s-s-T½)" Then
                VGc = 1 : VGd = 2 : VGee = 3 : VGf = 1 : VGg = 1 : VGh = 1 : VGi = 3 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_a4_10 = 10 : I_5_12 = 12 : I_b6_14 = 14 : I_6_15 = 15 : I_8_21 = 21
            ElseIf VGoo = "Mela Bhavapriya (s-T-T½-s-s-T-T)" Then
                VGc = 1 : VGd = 2 : VGee = 3 : VGf = 1 : VGg = 1 : VGh = 2 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_a4_10 = 10 : I_5_12 = 12 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Todi Theta (s-T-T½-s-s-T½-s)" OrElse VGoo = "Mela Subhapantuvarali (45) (s-T-T½-s-s-T½-s)" Then
                VGc = 1 : VGd = 2 : VGee = 3 : VGf = 1 : VGg = 1 : VGh = 3 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_a4_10 = 10 : I_5_12 = 12 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Todi Theta Scale", "Mela Subhapantuvarali (45)"})
            ElseIf VGoo = "Mela Sadvidhamargini (46) (s-T-T½-s-T-s-T)" Then
                VGc = 1 : VGd = 2 : VGee = 3 : VGf = 1 : VGg = 2 : VGh = 1 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_a4_10 = 10 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Mela Suvarnangi (s-T-T½-s-T-T-s)" Then
                VGc = 1 : VGd = 2 : VGee = 3 : VGf = 1 : VGg = 2 : VGh = 2 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_a4_10 = 10 : I_5_12 = 12 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Mela Divyamani (48) (s-T-T½-s-T½-s-s)" Then
                VGc = 1 : VGd = 2 : VGee = 3 : VGf = 1 : VGg = 3 : VGh = 1 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_a4_10 = 10 : I_5_12 = 12 : I_a6_16 = 16 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Vijayasri (s-T-T½-s-2T-s)" Then
                VGc = 1 : VGd = 2 : VGee = 3 : VGf = 1 : VGg = 4 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_a4_10 = 10 : I_5_12 = 12 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Bhavani hexatônica (s-T-T½-T-T-T)" Then
                VGc = 1 : VGd = 2 : VGee = 3 : VGf = 2 : VGg = 2 : VGh = 2 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_b5_11 = 11 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Raga Gurjari Todi (s-T-T½-T-T½-s)" Then
                VGc = 1 : VGd = 2 : VGee = 3 : VGf = 2 : VGg = 3 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_b5_11 = 11 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Chhaya Todi (s-T-T½-T-2T)" Then
                VGc = 1 : VGd = 2 : VGee = 3 : VGf = 2 : VGg = 4 : qtdeloop = 6
                I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_b5_11 = 11 : I_b6_14 = 14 : I_8_21 = 21
            ElseIf VGoo = "Raga Salagavarali (s-T-2T-T-s-T)" Then
                VGc = 1 : VGd = 2 : VGee = 4 : VGf = 2 : VGg = 1 : VGh = 2 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Pelog (s-T-2T-T½-T)" Then
                VGc = 1 : VGd = 2 : VGee = 4 : VGf = 3 : VGg = 2 : qtdeloop = 6
                I_1_0 = 0 : I_b2_2 = 2 : I_b3_5 = 5 : I_5_12 = 12 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Raga Bhatiyar (s-T½-s-s-s-T-T-s)" Then
                VGc = 1 : VGd = 3 : VGee = 1 : VGf = 1 : VGg = 1 : VGh = 2 : VGi = 2 : VGj = 1 : qtdeloop = 9
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_4_9 = 9 : I_b5_11 = 11 : I_5_12 = 12 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Enigmática de Verdi 1 (s-T½-s-s-T-T-s-s)" Then
                VGc = 1 : VGd = 3 : VGee = 1 : VGf = 1 : VGg = 2 : VGh = 2 : VGi = 1 : VGj = 1 : qtdeloop = 9
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_4_9 = 9 : I_b5_11 = 11 : I_b6_14 = 14 : I_b7_17 = 17 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Oriental 1 (s-T½-s-s-T-T-T)" Then
                VGc = 1 : VGd = 3 : VGee = 1 : VGf = 1 : VGg = 2 : VGh = 2 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_4_9 = 9 : I_b5_11 = 11 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Persa 1 (s-T½-s-s-T-T½-s)" OrElse VGoo = "Raga Lalita (s-T½-s-s-T-T½-s)" Then
                VGc = 1 : VGd = 3 : VGee = 1 : VGf = 1 : VGg = 2 : VGh = 3 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_4_9 = 9 : I_b5_11 = 11 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Persa 1", "Raga Lalita"})
            ElseIf VGoo = "Oriental 2 (s-T½-s-s-T½-s-T)" Then
                VGc = 1 : VGd = 3 : VGee = 1 : VGf = 1 : VGg = 3 : VGh = 1 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_4_9 = 9 : I_b5_11 = 11 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Lidia cromática (s-T½-s-s-T½-T-s)" Then
                VGc = 1 : VGd = 3 : VGee = 1 : VGf = 1 : VGg = 3 : VGh = 2 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_4_9 = 9 : I_b5_11 = 11 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Saurastra (s-T½-s-T-s-s-T-s)" Then
                VGc = 1 : VGd = 3 : VGee = 1 : VGf = 2 : VGg = 1 : VGh = 1 : VGi = 2 : VGj = 1 : qtdeloop = 9
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Mela Gayakapriya (13) (s-T½-s-T-s-s-T½)" OrElse VGoo = "Mela Mayamalavagaula (15) (s-T½-s-T-s-s-T½)" OrElse VGoo = "Cíngara hexatônica (s-T½-s-T-s-s-T½)" Then
                VGc = 1 : VGd = 3 : VGee = 1 : VGf = 2 : VGg = 1 : VGh = 1 : VGi = 3 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_6_15 = 15 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Mela Gayakapriya (13)", "Mela Mayamalavagaula (15)", "Cíngara hexatônica"})
            ElseIf VGoo = "Maqam Hijaz (s-T½-s-T-s-T-s-s)" Then
                VGc = 1 : VGd = 3 : VGee = 1 : VGf = 2 : VGg = 1 : VGh = 2 : VGi = 1 : VGj = 1 : qtdeloop = 9
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_b7_17 = 17 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Purna Pancama (s-T½-s-T-s-2T)" Then
                VGc = 1 : VGd = 3 : VGee = 1 : VGf = 2 : VGg = 1 : VGh = 4 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_8_21 = 21
            ElseIf VGoo = "Maqam Hicaz (s-T½-s-T-T-s-T)" OrElse VGoo = "Mela Chakravakam (16) (s-T½-s-T-T-s-T)" OrElse VGoo = "Raga Ahir Bhairav (s-T½-s-T-T-s-T)" Then
                VGc = 1 : VGd = 3 : VGee = 1 : VGf = 2 : VGg = 2 : VGh = 1 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Maqam Hicaz", "Mela Chakravakam (16)", "Raga Ahir Bhairav"})
            ElseIf VGoo = "Mela Suryakantam (17) (s-T½-s-T-T-T-s)" Then
                VGc = 1 : VGd = 3 : VGee = 1 : VGf = 2 : VGg = 2 : VGh = 2 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Kalavati (s-T½-s-T-T-T½)" Then
                VGc = 1 : VGd = 3 : VGee = 1 : VGf = 2 : VGg = 2 : VGh = 3 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_8_21 = 21
            ElseIf VGoo = "Mela Hatakambari (18) (s-T½-s-T-T½-s-s)" Then
                VGc = 1 : VGd = 3 : VGee = 1 : VGf = 2 : VGg = 3 : VGh = 1 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_a6_16 = 16 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Espanhola hexatônica (s-T½-s-T-T½-T)" Then
                VGc = 1 : VGd = 3 : VGee = 1 : VGf = 2 : VGg = 3 : VGh = 2 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Raga Gaula (s-T½-s-T-2T-s)" Then
                VGc = 1 : VGd = 3 : VGee = 1 : VGf = 2 : VGg = 4 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Simétrica hexatônica (s-T½-s-T½-s-T½)" Then
                VGc = 1 : VGd = 3 : VGee = 1 : VGf = 3 : VGg = 1 : VGh = 3 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_4_9 = 9 : I_a5_13 = 13 : I_6_15 = 15 : I_8_21 = 21
            ElseIf VGoo = "Enigmática de Verdi 2 (s-T½-s-T½-T-s-s)" Then
                VGc = 1 : VGd = 3 : VGee = 1 : VGf = 3 : VGg = 2 : VGh = 1 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_4_9 = 9 : I_a5_13 = 13 : I_a6_16 = 16 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Vasantabhairavi (s-T½-s-T½-T-T)" Then
                VGc = 1 : VGd = 3 : VGee = 1 : VGf = 3 : VGg = 2 : VGh = 2 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_4_9 = 9 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Raga Megharanjani (s-T½-s-T½-2T)" OrElse VGoo = "Siria (s-T½-s-T½-2T)" Then
                VGc = 1 : VGd = 3 : VGee = 1 : VGf = 3 : VGg = 4 : qtdeloop = 6
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_4_9 = 9 : I_b6_14 = 14 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Raga Megharanjani", "Síria"})
            ElseIf VGoo = "Raga Rudra Pancama (s-T½-s-2T-s-T)" Then
                VGc = 1 : VGd = 3 : VGee = 1 : VGf = 4 : VGg = 1 : VGh = 2 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_4_9 = 9 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Raga Vasanta (s-T½-s-2T-T-s)" Then
                VGc = 1 : VGd = 3 : VGee = 1 : VGf = 4 : VGg = 2 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_4_9 = 9 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Megharanji (s-T½-s-3T-s)" Then
                VGc = 1 : VGd = 3 : VGee = 1 : VGf = 6 : VGg = 1 : qtdeloop = 6
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_4_9 = 9 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Mela Dhavalambari (49) (s-T½-T-s-s-s-T½)" Then
                VGc = 1 : VGd = 3 : VGee = 2 : VGf = 1 : VGg = 1 : VGh = 1 : VGi = 3 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_a4_10 = 10 : I_5_12 = 12 : I_b6_14 = 14 : I_6_15 = 15 : I_8_21 = 21
            ElseIf VGoo = "Húngara Maior 2 (s-T½-T-s-s-T-T)" OrElse VGoo = "Mela Namanarayani (50) (s-T½-T-s-s-T-T)" OrElse VGoo = "Cíngara Maior 2 (s-T½-T-s-s-T-T)" Then
                VGc = 1 : VGd = 3 : VGee = 2 : VGf = 1 : VGg = 1 : VGh = 2 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_a4_10 = 10 : I_5_12 = 12 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Húngara Maior 2", "Mela Namanarayani (50)", "Cíngara Maior 2"})
            ElseIf VGoo = "Hipolidia cromática (s-T½-T-s-s-T½-s)" OrElse VGoo = "Mela Kamavardhani (51) (s-T½-T-s-s-T½-s)" OrElse VGoo = "Purvi Theta (s-T½-T-s-s-T½-s)" OrElse VGoo = "Raga Shri (s-T½-T-s-s-T½-s)" Then
                VGc = 1 : VGd = 3 : VGee = 2 : VGf = 1 : VGg = 1 : VGh = 3 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_a4_10 = 10 : I_5_12 = 12 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Hipolidia cromática", "Mela Kamavardhani (51)", "Purvi Theta", "Raga Shri"})
            ElseIf VGoo = "Raga Dhavalangam (s-T½-T-s-s-2T)" Then
                VGc = 1 : VGd = 3 : VGee = 2 : VGf = 1 : VGg = 1 : VGh = 4 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_a4_10 = 10 : I_5_12 = 12 : I_b6_14 = 14 : I_8_21 = 21
            ElseIf VGoo = "Mela Ramapriya (52) (s-T½-T-s-T-s-T)" Then
                VGc = 1 : VGd = 3 : VGee = 2 : VGf = 1 : VGg = 2 : VGh = 1 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_a4_10 = 10 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Peiraiotikos (s-T½-T-s-T-T-s)" OrElse VGoo = "Mela Gamanasrama (53) (s-T½-T-s-T-T-s)" OrElse VGoo = "Theta, Marva (s-T½-T-s-T-T-s)" Then
                VGc = 1 : VGd = 3 : VGee = 2 : VGf = 1 : VGg = 2 : VGh = 2 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_a4_10 = 10 : I_5_12 = 12 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Peiraiotikos", "Mela Gamanasrama (53)", "Theta, Marva"})
            ElseIf VGoo = "Mela Visvambhari (s-T½-T-s-T½-s-s)" Then
                VGc = 1 : VGd = 3 : VGee = 2 : VGf = 1 : VGg = 3 : VGh = 1 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_a4_10 = 10 : I_5_12 = 12 : I_a6_16 = 16 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Mandari (s-T½-T-s-2T-s)" Then
                VGc = 1 : VGd = 3 : VGee = 2 : VGf = 1 : VGg = 4 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_a4_10 = 10 : I_5_12 = 12 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Hejjajji (s-T½-T-T-s-T½)" Then
                VGc = 1 : VGd = 3 : VGee = 2 : VGf = 2 : VGg = 1 : VGh = 3 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_a4_10 = 10 : I_a5_13 = 13 : I_6_15 = 15 : I_8_21 = 21
            ElseIf VGoo = "Enigmática de Verdi 3 (s-T½-T-T-T-s-s)" Then
                VGc = 1 : VGd = 3 : VGee = 2 : VGf = 2 : VGg = 2 : VGh = 1 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_a4_10 = 10 : I_a5_13 = 13 : I_a6_16 = 16 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Enigmática 2 (s-T½-T-T-T½-s)" Then
                VGc = 1 : VGd = 3 : VGee = 2 : VGf = 2 : VGg = 3 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_a4_10 = 10 : I_a5_13 = 13 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Napolitana 1 (s-T½-T-T½-s-T)" OrElse VGoo = "Prometheus napolitana (s-T½-T-T½-s-T)" Then
                VGc = 1 : VGd = 3 : VGee = 2 : VGf = 3 : VGg = 1 : VGh = 2 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_a4_10 = 10 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Napolitana 1", "Prometheus napolitana"})
            ElseIf VGoo = "Raga Hamsanandi (s-T½-T-T½-T-s)" Then
                VGc = 1 : VGd = 3 : VGee = 2 : VGf = 3 : VGg = 2 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_b5_11 = 11 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Kalagada (s-T½-T½-s-s-T½)" Then
                VGc = 1 : VGd = 3 : VGee = 3 : VGf = 1 : VGg = 1 : VGh = 3 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_5_12 = 12 : I_b6_14 = 14 : I_6_15 = 15 : I_8_21 = 21
            ElseIf VGoo = "Raga Bauli (s-T½-T½-s-T½-s)" Then
                VGc = 1 : VGd = 3 : VGee = 3 : VGf = 1 : VGg = 3 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_5_12 = 12 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Reva (s-T½-T½-s-2T)" Then
                VGc = 1 : VGd = 3 : VGee = 3 : VGf = 1 : VGg = 4 : qtdeloop = 6
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_5_12 = 12 : I_b6_14 = 14 : I_8_21 = 21
            ElseIf VGoo = "Raga Malayamarutam (s-T½-T½-T-s-T)" Then
                VGc = 1 : VGd = 3 : VGee = 3 : VGf = 2 : VGg = 1 : VGh = 2 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Pentatônica Alt.b2 (s-T½-T½-T-T½)" OrElse VGoo = "Raga Rasika Ranjani (s-T½-T½-T-T½)" OrElse VGoo = "Skriabin 2 (s-T½-T½-T-T½)" Then
                VGc = 1 : VGd = 3 : VGee = 3 : VGf = 2 : VGg = 3 : qtdeloop = 6
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_5_12 = 12 : I_6_15 = 15 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Pentatônica Alt.b2", "Raga Rasika Ranjani", "Skriabin 2"})
            ElseIf VGoo = "Raga Manaranjani I (s-T½-T½-T½-T)" Then
                VGc = 1 : VGd = 3 : VGee = 3 : VGf = 3 : VGg = 2 : qtdeloop = 6
                I_1_0 = 0 : I_b2_2 = 2 : I_3_6 = 6 : I_5_12 = 12 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Iwato (s-2T-s-2T-T)" Then
                VGc = 1 : VGd = 4 : VGee = 1 : VGf = 4 : VGg = 2 : qtdeloop = 6
                I_1_0 = 0 : I_b2_2 = 2 : I_4_9 = 9 : I_b5_11 = 11 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Raga Kalakanthi (s-2T-T-s-s-T½)" Then
                VGc = 1 : VGd = 4 : VGee = 2 : VGf = 1 : VGg = 1 : VGh = 3 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_6_15 = 15 : I_8_21 = 21
            ElseIf VGoo = "Niagari hexatônica (s-2T-T-s-T-T)" OrElse VGoo = "Raga Phenadyuti (s-2T-T-s-T-T)" Then
                VGc = 1 : VGd = 4 : VGee = 2 : VGf = 1 : VGg = 2 : VGh = 2 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Niagari hexatónica", "Raga Phenadyuti"})
            ElseIf VGoo = "Raga Padi (s-2T-T-s-T½-s)" Then
                VGc = 1 : VGd = 4 : VGee = 2 : VGf = 1 : VGg = 3 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Hon-kumoi-joshi (s-2T-T-s-2T)" OrElse VGoo = "Japonesa (A) (s-2T-T-s-2T)" OrElse VGoo = "Raga Salanganata (s-2T-T-s-2T)" Then
                VGc = 1 : VGd = 4 : VGee = 2 : VGf = 1 : VGg = 4 : qtdeloop = 6
                I_1_0 = 0 : I_b2_2 = 2 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Hon-kumoi-joshi", "Japonesa (A)", "Raga Salanganata"})
            ElseIf VGoo = "Raga Rasavali (s-2T-T-T-s-T)" Then
                VGc = 1 : VGd = 4 : VGee = 2 : VGf = 2 : VGg = 1 : VGh = 2 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Raga Jivantika (s-2T-T-T-T-s)" Then
                VGc = 1 : VGd = 4 : VGee = 2 : VGf = 2 : VGg = 2 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_b2_2 = 2 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Pentatônica Neutra 2 (s-2T-T-T-T½)" OrElse VGoo = "Raga Manaranjani II (s-2T-T-T-T½)" Then
                VGc = 1 : VGd = 4 : VGee = 2 : VGf = 2 : VGg = 3 : qtdeloop = 6
                I_1_0 = 0 : I_b2_2 = 2 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Pentatônica neutra 2", "Raga Manaranjani II"})
            ElseIf VGoo = "Kokin-joshi (s-2T-T-T½-T)" OrElse VGoo = "Raga Vibhavari (s-2T-T-T½-T)" Then
                VGc = 1 : VGd = 4 : VGee = 2 : VGf = 3 : VGg = 2 : qtdeloop = 6
                I_1_0 = 0 : I_b2_2 = 2 : I_4_9 = 9 : I_5_12 = 12 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Kokin-joshi", "Raga Vibhavari"})
            ElseIf VGoo = "Raga Gauri (s-2T-T-2T-s)" Then
                VGc = 1 : VGd = 4 : VGee = 2 : VGf = 4 : VGg = 1 : qtdeloop = 6
                I_1_0 = 0 : I_b2_2 = 2 : I_4_9 = 9 : I_5_12 = 12 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Kshanika (s-2T-T½-T½-s)" Then
                VGc = 1 : VGd = 4 : VGee = 3 : VGf = 3 : VGg = 1 : qtdeloop = 6
                I_1_0 = 0 : I_b2_2 = 2 : I_4_9 = 9 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Saugandhini (s-2T½-s-s-2T)" Then
                VGc = 1 : VGd = 5 : VGee = 1 : VGf = 1 : VGg = 4 : qtdeloop = 6
                I_1_0 = 0 : I_b2_2 = 2 : I_a4_10 = 10 : I_5_12 = 12 : I_b6_14 = 14 : I_8_21 = 21
            ElseIf VGoo = "Raga Deshgaur (s-3T-s-T½-s)" Then
                VGc = 1 : VGd = 6 : VGee = 1 : VGf = 3 : VGg = 1 : qtdeloop = 6
                I_1_0 = 0 : I_b2_2 = 2 : I_5_12 = 12 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Lavangi (s-3T-T½-T)" Then
                VGc = 1 : VGd = 6 : VGee = 3 : VGf = 2 : qtdeloop = 5
                I_1_0 = 0 : I_b2_2 = 2 : I_5_12 = 12 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Be-Bop menor (T-s-s-s-T-T-s-T)" Then
                VGc = 2 : VGd = 1 : VGee = 1 : VGf = 1 : VGg = 2 : VGh = 2 : VGi = 1 : VGj = 2 : qtdeloop = 9
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Nonatônica (T-s-s-T-s-s-s-T-s)" Then
                VGc = 2 : VGd = 1 : VGee = 1 : VGf = 2 : VGg = 1 : VGh = 1 : VGi = 1 : VGj = 2 : VGk = 1 : qtdeloop = 10
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_3_6 = 6 : I_b5_11 = 11 : I_5_12 = 12 : I_a5_13 = 13 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Ramdasi Malhar (T-s-s-T-T-s-s-s-s)" Then
                VGc = 2 : VGd = 1 : VGee = 1 : VGf = 2 : VGg = 2 : VGh = 1 : VGi = 1 : VGj = 1 : VGk = 1 : qtdeloop = 10
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_3_6 = 6 : I_a4_10 = 10 : I_a5_13 = 13 : I_6_15 = 15 : I_b7_17 = 17 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Hipodórica cromática (T-s-s-T½-s-s-T½)" Then
                VGc = 2 : VGd = 1 : VGee = 1 : VGf = 3 : VGg = 1 : VGh = 1 : VGi = 3 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_3_6 = 6 : I_5_12 = 12 : I_b6_14 = 14 : I_6_15 = 15 : I_8_21 = 21
            ElseIf VGoo = "Algeriana (T-s-T-s-s-s-T½-s)" Then
                VGc = 2 : VGd = 1 : VGee = 2 : VGf = 1 : VGg = 1 : VGh = 1 : VGi = 3 : VGj = 1 : qtdeloop = 9
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_4_9 = 9 : I_b5_11 = 11 : I_5_12 = 12 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Árabe 2 (T-s-T-s-T-s-T-s)" OrElse VGoo = "Octatônica (T-s) (T-s-T-s-T-s-T-s)" OrElse VGoo = "Auxiliar Diminuída (T-s-T-s-T-s-T-s)" OrElse VGoo = "Diminuída 1 (T-s-T-s-T-s-T-s)" OrElse VGoo = "Simétrica 1 (T-s-T-s-T-s-T-s)" OrElse VGoo = "Tom-semitom (T-s-T-s-T-s-T-s)" Then
                VGc = 2 : VGd = 1 : VGee = 2 : VGf = 1 : VGg = 2 : VGh = 1 : VGi = 2 : VGj = 1 : qtdeloop = 9
                If VGoo = "Tom-semitom (T-s-T-s-T-s-T-s)" Then I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_4_9 = 9 : I_b5_11 = 11 : I_b6_14 = 14 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
                If VGoo = "Simétrica 1 (T-s-T-s-T-s-T-s)" Then I_1_0 = 0 : I_2_3 = 3 : I_a2_4 = 4 : I_4_9 = 9 : I_a4_10 = 10 : I_a5_13 = 13 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
                If VGoo = "Diminuída 1 (T-s-T-s-T-s-T-s)" OrElse VGoo = "Auxiliar Diminuída (T-s-T-s-T-s-T-s)" OrElse VGoo = "Octatônica (T-s) (T-s-T-s-T-s-T-s)" OrElse VGoo = "Árabe 2 (T-s-T-s-T-s-T-s)" Then I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_4_9 = 9 : I_a4_10 = 10 : I_a5_13 = 13 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Árabe 2", "Octatônica (T-s)", "Auxiliar Diminuída", "Diminuída 1", "Simétrica 1", "Tom-Semitom"})
            ElseIf VGoo = "Dórica alterada (T-s-T-s-T-T-T)" OrElse VGoo = "Semi-diminuída #2 (Lócrio #2) (T-s-T-s-T-T-T)" OrElse VGoo = "Semidiminuída (T-s-T-s-T-T-T)" Then
                VGc = 2 : VGd = 1 : VGee = 2 : VGf = 1 : VGg = 2 : VGh = 2 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_4_9 = 9 : I_b5_11 = 11 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Dórica alterada", "Semi-diminuída", "Semi-diminuída #2 (Lócrio #2)"})
            ElseIf VGoo = "Árabe 3 (T-s-T-s-T½-s-T)" OrElse VGoo = "Maqam Karcigar (T-s-T-s-T½-s-T)" Then
                VGc = 2 : VGd = 1 : VGee = 2 : VGf = 1 : VGg = 3 : VGh = 1 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_4_9 = 9 : I_b5_11 = 11 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Árabe 3", "Maqam Karcigar"})
            ElseIf VGoo = "Hexatônica Piramidal (T-s-T-s-T½-T½)" Then
                VGc = 2 : VGd = 1 : VGee = 2 : VGf = 1 : VGg = 3 : VGh = 3 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_4_9 = 9 : I_b5_11 = 11 : I_6_15 = 15 : I_8_21 = 21
            ElseIf VGoo = "Zirafkend T-s-T-T-s-" Then
                VGc = 2 : VGd = 1 : VGee = 2 : VGf = 2 : VGg = 1 : VGh = 1 : qtdeloop = 9
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_6_15 = 15 : I_8_21 = 21
            ElseIf VGoo = "Raga Pilu (T-s-T-T-s-s-s-s-s)" Then
                VGc = 2 : VGd = 1 : VGee = 2 : VGf = 2 : VGg = 1 : VGh = 1 : VGi = 1 : VGj = 1 : VGk = 1 : qtdeloop = 10
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_6_15 = 15 : I_b7_17 = 17 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Gregoriana (T-s-T-T-s-s-s-T)" OrElse VGoo = "Raga Mukhari (T-s-T-T-s-s-s-T)" Then
                VGc = 2 : VGd = 1 : VGee = 2 : VGf = 2 : VGg = 1 : VGh = 1 : VGi = 1 : VGj = 2 : qtdeloop = 9
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Gregorian", "Raga Mukhari"})
            ElseIf VGoo = "Maqam Nahawand (T-s-T-T-s-T-s-s)" Then
                VGc = 2 : VGd = 1 : VGee = 2 : VGf = 2 : VGg = 1 : VGh = 2 : VGi = 1 : VGj = 1 : qtdeloop = 9
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_b7_17 = 17 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Árabe 4 (T-s-T-T-T-s-s-s)" OrElse VGoo = "Raga Mian Ki Malhar (T-s-T-T-T-s-s-s)" Then
                VGc = 2 : VGd = 1 : VGee = 2 : VGf = 2 : VGg = 2 : VGh = 1 : VGi = 1 : VGj = 1 : qtdeloop = 9
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Árabe 4", "Raga Mian Ki Malhar"})
            ElseIf VGoo = "Mela Varunapriya (24) (T-s-T-T-T½-s-s)" Then
                VGc = 2 : VGd = 1 : VGee = 2 : VGf = 2 : VGg = 3 : VGh = 1 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_4_9 = 9 : I_5_12 = 12 : I_a6_16 = 16 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Mela Syamalangi (55) (T-s-T½-s-s-s-T½)" Then
                VGc = 2 : VGd = 1 : VGee = 3 : VGf = 1 : VGg = 1 : VGh = 1 : VGi = 3 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_a4_10 = 10 : I_5_12 = 12 : I_b6_14 = 14 : I_6_15 = 15 : I_8_21 = 21
            ElseIf VGoo = "Húngara Menor 1 (T-s-T½-s-s-T-T)" OrElse VGoo = "Maqam Suzdil (T-s-T½-s-s-T-T)" OrElse VGoo = "Mela Sanmukhapriya (56) (T-s-T½-s-s-T-T)" Then
                VGc = 2 : VGd = 1 : VGee = 3 : VGf = 1 : VGg = 1 : VGh = 2 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_a4_10 = 10 : I_5_12 = 12 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Húngara menor 1", "Maqam Suzdil", "Mela Sanmukhapriya (56)"})
            ElseIf VGoo = "Húngara Menor 2 (T-s-T½-s-s-T½-s)" OrElse VGoo = "Hungarian Gypsy (T-s-T½-s-s-T½-s)" OrElse VGoo = "Mela Simhendramadhyama (57) (T-s-T½-s-s-T½-s)" OrElse VGoo = "Niavent (T-s-T½-s-s-T½-s)" OrElse VGoo = "Cíngara Menor (T-s-T½-s-s-T½-s)" Then
                VGc = 2 : VGd = 1 : VGee = 3 : VGf = 1 : VGg = 1 : VGh = 3 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_a4_10 = 10 : I_5_12 = 12 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Húngara menor 2", "Húngara Cigana", "Mela Simhendramadhyama (57)", "Niavent", "Cíngara Menor"})
            ElseIf VGoo = "Hedjaz (T-s-T½-s-T-s-T)" OrElse VGoo = "Maqam Nakriz (T-s-T½-s-T-s-T)" OrElse VGoo = "Mela Hemavati (58) (T-s-T½-s-T-s-T)" OrElse VGoo = "Romana menor (T-s-T½-s-T-s-T)" OrElse VGoo = "Souzinak (T-s-T½-s-T-s-T)" Then
                VGc = 2 : VGd = 1 : VGee = 3 : VGf = 1 : VGg = 2 : VGh = 1 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_a4_10 = 10 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Hedjaz", "Maqam Nakriz", "Mela Hemavati (58)", "Romana", "Souzinak"})
            ElseIf VGoo = "Lidia diminuída (T-s-T½-s-T-T-s)" OrElse VGoo = "Mela Dharmavati (59) (T-s-T½-s-T-T-s)" Then
                VGc = 2 : VGd = 1 : VGee = 3 : VGf = 1 : VGg = 2 : VGh = 2 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_a4_10 = 10 : I_5_12 = 12 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Lídia diminuída", "Mela Dharmavati (59)"})
            ElseIf VGoo = "Mela Nitimati (60) (T-s-T½-s-T½-s-s)" Then
                VGc = 2 : VGd = 1 : VGee = 3 : VGf = 1 : VGg = 3 : VGh = 1 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_a4_10 = 10 : I_5_12 = 12 : I_a6_16 = 16 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Hira-joshi (T-s-2T-s-2T)" OrElse VGoo = "Pentatônica Alt.b3/b6 (T-s-2T-s-2T)" Then
                VGc = 2 : VGd = 1 : VGee = 4 : VGf = 1 : VGg = 4 : qtdeloop = 6
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_5_12 = 12 : I_b6_14 = 14 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Hira-joshi", "Pentatônica Alt.b3/b6"})
            ElseIf VGoo = "Hawayana 1 (T-s-2T-T-T-s)" Then
                VGc = 2 : VGd = 1 : VGee = 4 : VGf = 2 : VGg = 2 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_5_12 = 12 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Pentatônica Menor 4 (T-s-2T-T½-T)" Then
                VGc = 2 : VGd = 1 : VGee = 4 : VGf = 3 : VGg = 2 : qtdeloop = 6
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_5_12 = 12 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Sambah (T-s-s-T½-s-T-T)" Then
                VGc = 2 : VGd = 1 : VGee = 1 : VGf = 3 : VGg = 1 : VGh = 2 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_b4_8 = 8 : I_5_12 = 12 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Warao tetratônica (T-s-3T½-T)" Then
                VGc = 2 : VGd = 1 : VGee = 7 : VGf = 2 : qtdeloop = 5
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Taishikicho (T-T-s-s-s-T-s-s-s)" Then
                VGc = 2 : VGd = 2 : VGee = 1 : VGf = 1 : VGg = 1 : VGh = 2 : VGi = 1 : VGj = 1 : VGk = 1 : qtdeloop = 10
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_4_9 = 9 : I_b5_11 = 11 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Genus diatonicum veterum (T-T-s-s-s-T-T-s)" OrElse VGoo = "Ishikotsucho ou Ichikosucho (T-T-s-s-s-T-T-s)" Then
                VGc = 2 : VGd = 2 : VGee = 1 : VGf = 1 : VGg = 1 : VGh = 2 : VGi = 2 : VGj = 1 : qtdeloop = 9
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_4_9 = 9 : I_b5_11 = 11 : I_5_12 = 12 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Genus diatonicum veterum", "Ichikosucho", "Ishikotsucho"})
            ElseIf VGoo = "Raga Dipak (T-T-s-s-s-2T½)" Then
                VGc = 2 : VGd = 2 : VGee = 1 : VGf = 1 : VGg = 1 : VGh = 5 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_4_9 = 9 : I_b5_11 = 11 : I_5_12 = 12 : I_8_21 = 21
            ElseIf VGoo = "Árabe 1 (T-T-s-s-T-T-T)" OrElse VGoo = "Lócrio Maior (T-T-s-s-T-T-T)" Then
                VGc = 2 : VGd = 2 : VGee = 1 : VGf = 1 : VGg = 2 : VGh = 2 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_4_9 = 9 : I_b5_11 = 11 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Árabe 1", "Lócrio Maior"})
            ElseIf VGoo = "Be-Bop Maior (T-T-s-T-s-s-T-s)" Then
                VGc = 2 : VGd = 2 : VGee = 1 : VGf = 2 : VGg = 1 : VGh = 1 : VGi = 2 : VGj = 1 : qtdeloop = 9
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Mela Carukesi (T-T-s-T-s-s-T½)" OrElse VGoo = "Mela Mararanjani (25) (T-T-s-T-s-s-T½)" Then
                VGc = 2 : VGd = 2 : VGee = 1 : VGf = 2 : VGg = 1 : VGh = 1 : VGi = 3 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_6_15 = 15 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Mela Carukesi", "Mela Mararanjani (25)"})
            ElseIf VGoo = "Hindu, Hindustan ou Indostán (T-T-s-T-s-T-T)" OrElse VGoo = "Mela Charukesi (26) (T-T-s-T-s-T-T)" OrElse VGoo = "Mischung 6 (T-T-s-T-s-T-T)" Then
                VGc = 2 : VGd = 2 : VGee = 1 : VGf = 2 : VGg = 1 : VGh = 2 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Hindu", "Hindustan", "Mela Charukesi (26)", "Mischung 6"})
            ElseIf VGoo = "Etíope 3 (T-T-s-T-s-T½-s)" OrElse VGoo = "Mela Sarasangi (27) (T-T-s-T-s-T½-s)" OrElse VGoo = "Harmônica Maior (T-T-s-T-s-T½-s)" OrElse VGoo = "Mischung 2 (T-T-s-T-s-T½-s)" Then
                VGc = 2 : VGd = 2 : VGee = 1 : VGf = 2 : VGg = 1 : VGh = 3 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Etíope 3", "Mela Sarasangi (27)", "Harmônica Maior", "Mischung 2"})
            ElseIf VGoo = "Be-Bop Dominante (T-T-s-T-T-s-s-s)" OrElse VGoo = "Genus diatonicum (T-T-s-T-T-s-s-s)" OrElse VGoo = "Maqam Shawq (T-T-s-T-T-s-s-s)" OrElse VGoo = "Rast (T-T-s-T-T-s-s-s)" OrElse VGoo = "China octatônica (T-T-s-T-T-s-s-s)" Then
                VGc = 2 : VGd = 2 : VGee = 1 : VGf = 2 : VGg = 2 : VGh = 1 : VGi = 1 : VGj = 1 : qtdeloop = 9
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Be-Bop Dominante", "Genus diatonicum", "Maqam Shawq", "Rast", "China octatônica"})
            ElseIf VGoo = "Raga Kambhoji (T-T-s-T-T-T½)" Then
                VGc = 2 : VGd = 2 : VGee = 1 : VGf = 2 : VGg = 2 : VGh = 3 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_8_21 = 21
            ElseIf VGoo = "Mela Naganandini (30) (T-T-s-T-T½-s-s)" Then
                VGc = 2 : VGd = 2 : VGee = 1 : VGf = 2 : VGg = 3 : VGh = 1 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_a6_16 = 16 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Siva Kambhoji (T-T-s-T-T½-T)" Then
                VGc = 2 : VGd = 2 : VGee = 1 : VGf = 2 : VGg = 3 : VGh = 2 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Raga Nalinakanti (T-T-s-T-2T-s)" Then
                VGc = 2 : VGd = 2 : VGee = 1 : VGf = 2 : VGg = 4 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Mixolidia aumentada (T-T-s-T½-s-s-T)" Then
                VGc = 2 : VGd = 2 : VGee = 1 : VGf = 3 : VGg = 1 : VGh = 1 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_4_9 = 9 : I_a5_13 = 13 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Judaica (T-T-s-T½-s-T-s)" OrElse VGoo = "Jônio aumentada (T-T-s-T½-s-T-s)" Then
                VGc = 2 : VGd = 2 : VGee = 1 : VGf = 3 : VGg = 1 : VGh = 2 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_4_9 = 9 : I_a5_13 = 13 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Judaica", "Jônio aumentada"})
            ElseIf VGoo = "Raga Sarasanana (T-T-s-T½-T½-s)" Then
                VGc = 2 : VGd = 2 : VGee = 1 : VGf = 3 : VGg = 3 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_4_9 = 9 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Ragesri (T-T-s-2T-s-s-s)" Then
                VGc = 2 : VGd = 2 : VGee = 1 : VGf = 4 : VGg = 1 : VGh = 1 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_4_9 = 9 : I_6_15 = 15 : I_b7_17 = 17 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Rageshri (T-T-s-2T-s-T)" Then
                VGc = 2 : VGd = 2 : VGee = 1 : VGf = 4 : VGg = 1 : VGh = 2 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_4_9 = 9 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Raga Hamsa Vinodini (T-T-s-2T-T-s)" Then
                VGc = 2 : VGd = 2 : VGee = 1 : VGf = 4 : VGg = 2 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_4_9 = 9 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Mela Kantamani (61) (T-T-T-s-s-s-T½)" Then
                VGc = 2 : VGd = 2 : VGee = 2 : VGf = 1 : VGg = 1 : VGh = 1 : VGi = 3 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_a4_10 = 10 : I_5_12 = 12 : I_b6_14 = 14 : I_6_15 = 15 : I_8_21 = 21
            ElseIf VGoo = "Mela Latangi (63) (T-T-T-s-s-T½-s)" Then
                VGc = 2 : VGd = 2 : VGee = 2 : VGf = 1 : VGg = 1 : VGh = 3 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_a4_10 = 10 : I_5_12 = 12 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Overtone (T-T-T-s-T-s-T)" OrElse VGoo = "Overtone Dominant (T-T-T-s-T-s-T)" OrElse VGoo = "Mela Vaschaspati (64) (T-T-T-s-T-s-T)" OrElse VGoo = "Lidia b7 (T-T-T-s-T-s-T)" Then
                VGc = 2 : VGd = 2 : VGee = 2 : VGf = 1 : VGg = 2 : VGh = 1 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_a4_10 = 10 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Overtone", "Overtone Dominante", "Mela Vaschaspati (64)", "Lidia b7"})
            ElseIf VGoo = "Raga Yamuna Kalyani (T-T-T-s-T-T½)" OrElse VGoo = "China antigua (T-T-T-s-T-T½)" Then
                VGc = 2 : VGd = 2 : VGee = 2 : VGf = 1 : VGg = 2 : VGh = 3 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_a4_10 = 10 : I_5_12 = 12 : I_6_15 = 15 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Raga Yamuna Kalyani", "China antigua"})
            ElseIf VGoo = "Mela Chitrambari (66) (T-T-T-s-T½-s-s)" Then
                VGc = 2 : VGd = 2 : VGee = 2 : VGf = 1 : VGg = 3 : VGh = 1 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_a4_10 = 10 : I_5_12 = 12 : I_a6_16 = 16 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Caturangini (T-T-T-s-2T-s)" Then
                VGc = 2 : VGd = 2 : VGee = 2 : VGf = 1 : VGg = 4 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_a4_10 = 10 : I_5_12 = 12 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Lidia aumentada (T-T-T-T-s-T-s)" Then
                VGc = 2 : VGd = 2 : VGee = 2 : VGf = 2 : VGg = 1 : VGh = 2 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_a4_10 = 10 : I_a5_13 = 13 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Esquimal Hexatônica 1 (T-T-T-T-T½-s)" Then
                VGc = 2 : VGd = 2 : VGee = 2 : VGf = 2 : VGg = 3 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_b5_11 = 11 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Prometheus (T-T-T-T½-s-T)" OrElse VGoo = "Raga Barbara (T-T-T-T½-s-T)" OrElse VGoo = "Skriabin 1 (T-T-T-T½-s-T)" Then
                VGc = 2 : VGd = 2 : VGee = 2 : VGf = 3 : VGg = 1 : VGh = 2 : qtdeloop = 7
                If VGoo = "Prometheus (T-T-T-T½-s-T)" Then I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_b5_11 = 11 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
                If VGoo = "Raga Barbara (T-T-T-T½-s-T)" OrElse VGoo = "Skriabin 1 (T-T-T-T½-s-T)" Then I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_a4_10 = 10 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Prometheus", "Raga Barbaraa", "Skriabin 1"})
            ElseIf VGoo = "Raga Mruganandana (T-T-T-T½-T-s)" Then
                VGc = 2 : VGd = 2 : VGee = 2 : VGf = 3 : VGg = 2 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_b5_11 = 11 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Kung (T-T-T-T½-T½)" OrElse VGoo = "Pentatônica Alt. b5 (T-T-T-T½-T½)" Then
                VGc = 2 : VGd = 2 : VGee = 2 : VGf = 3 : VGg = 3 : qtdeloop = 6
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_b5_11 = 11 : I_6_15 = 15 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Kung", "Pentatônica Alt. B5"})
            ElseIf VGoo = "Raga Kumurdaki (T-T-T-2T½-s)" Then
                VGc = 2 : VGd = 2 : VGee = 2 : VGf = 5 : VGg = 1 : qtdeloop = 6
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_b5_11 = 11 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Latika (T-T-T½-s-T½-s)" Then
                VGc = 2 : VGd = 2 : VGee = 3 : VGf = 1 : VGg = 3 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_5_12 = 12 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Pentatônica Alt. b6 (T-T-T½-s-2T)" OrElse VGoo = "Raga Bhupeshwari (T-T-T½-s-2T)" Then
                VGc = 2 : VGd = 2 : VGee = 3 : VGf = 1 : VGg = 4 : qtdeloop = 6
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_5_12 = 12 : I_b6_14 = 14 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Raga Bhupeshwari", "Pentatônica Alt. B6"})
            ElseIf VGoo = "Lidia hexatônica (T-T-T½-T-T-s)" OrElse VGoo = "Raga Kumud (T-T-T½-T-T-s)" Then
                VGc = 2 : VGd = 2 : VGee = 3 : VGf = 2 : VGg = 2 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_5_12 = 12 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Lidia hexatônica", "Raga Kumud"})
            ElseIf VGoo = "Pentatônica de Dominante (T-T-T½-T½-T)" OrElse VGoo = "Raga Hamsadhvani (T-T-T½-T½-T)" OrElse VGoo = "Shang (T-T-T½-T½-T)" Then
                VGc = 2 : VGd = 2 : VGee = 3 : VGf = 3 : VGg = 2 : qtdeloop = 6
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_5_12 = 12 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Pentatônica de Dominante", "Raga Hamsadhvani", "Shang"})
            ElseIf VGoo = "Pentatônica Maior 2 (T-T-T½-2T-s)" Then
                VGc = 2 : VGd = 2 : VGee = 3 : VGf = 4 : VGg = 1 : qtdeloop = 6
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_5_12 = 12 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Esquimal Tetratônica (T-T-T½-2T½)" Then
                VGc = 2 : VGd = 2 : VGee = 3 : VGf = 5 : qtdeloop = 5
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_5_12 = 12 : I_8_21 = 21
            ElseIf VGoo = "Raga Neroshta (T-T-2T-T-T)" Then
                VGc = 2 : VGd = 2 : VGee = 4 : VGf = 2 : VGg = 2 : qtdeloop = 6
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Nohkan (T-T½-s-T-s-T-s)" Then
                VGc = 2 : VGd = 3 : VGee = 1 : VGf = 2 : VGg = 1 : VGh = 2 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_4_9 = 9 : I_b5_11 = 11 : I_a5_13 = 13 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Navamanohari (T-T½-T-s-T-T)" Then
                VGc = 2 : VGd = 3 : VGee = 2 : VGf = 1 : VGg = 2 : VGh = 2 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Raga Bhinna Pancama (T-T½-T-s-T½-s)" Then
                VGc = 2 : VGd = 3 : VGee = 2 : VGf = 1 : VGg = 3 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Han-Kumoi (T-T½-T-s-2T)" OrElse VGoo = "Japonesa (B)(T-T½-T-s-2T)" OrElse VGoo = "Raga Shobhavari (T-T½-T-s-2T)" Then
                VGc = 2 : VGd = 3 : VGee = 2 : VGf = 1 : VGg = 4 : qtdeloop = 6
                I_1_0 = 0 : I_2_3 = 3 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Han-Kumoi", "Japonesa (B)", "Raga Shobhavari"})
            ElseIf VGoo = "Raga Sorati (T-T½-T-T-s-s-s)" Then
                VGc = 2 : VGd = 3 : VGee = 2 : VGf = 2 : VGg = 1 : VGh = 1 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Coreana 2 (T-T½-T-T-s-T)" OrElse VGoo = "Dom. sus 4 (T-T½-T-T-s-T)" OrElse VGoo = "Mixolidia hexatônica (T-T½-T-T-s-T)" OrElse VGoo = "Pyongjo (T-T½-T-T-s-T)" OrElse VGoo = "Raga Darbar (T-T½-T-T-s-T)" OrElse VGoo = "Yosen (T-T½-T-T-s-T)" Then
                VGc = 2 : VGd = 3 : VGee = 2 : VGf = 2 : VGg = 1 : VGh = 2 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Coreana 2", "Dominante sus 4", "Mixolidia hexatônica", "Pyongjo", "Raga Darbar", "Yosen"})
            ElseIf VGoo = "Raga Nagagandhari (T-T½-T-T-T-s)" Then
                VGc = 2 : VGd = 3 : VGee = 2 : VGf = 2 : VGg = 2 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Devakriya (T-T½-T-T-T½)" OrElse VGoo = "Ritusen (T-T½-T-T-T½)" OrElse VGoo = "Ujo (T-T½-T-T-T½)" OrElse VGoo = "Zhi (T-T½-T-T-T½)" Then
                VGc = 2 : VGd = 3 : VGee = 2 : VGf = 2 : VGg = 3 : qtdeloop = 6
                I_1_0 = 0 : I_2_3 = 3 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Raga Devakriya", "Ritusen", "Ujo", "Zhi"})
            ElseIf VGoo = "Raga Brindabani Sarang (T-T½-T-T½-s-s)" OrElse VGoo = "Raga Sarang (T-T½-T-T½-s-s)" Then
                VGc = 2 : VGd = 3 : VGee = 2 : VGf = 3 : VGg = 1 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_4_9 = 9 : I_5_12 = 12 : I_a6_16 = 16 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Raga Brindabani Sarang", "Raga Sarang"})
            ElseIf VGoo = "Raga Desh (T-T½-T-2T-s)" Then
                VGc = 2 : VGd = 3 : VGee = 2 : VGf = 4 : VGg = 1 : qtdeloop = 6
                I_1_0 = 0 : I_2_3 = 3 : I_4_9 = 9 : I_5_12 = 12 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Genus primum (T-T½-T-2T½)" Then
                VGc = 2 : VGd = 3 : VGee = 2 : VGf = 5 : qtdeloop = 5
                I_1_0 = 0 : I_2_3 = 3 : I_4_9 = 9 : I_5_12 = 12 : I_8_21 = 21
            ElseIf VGoo = "Chaio (T-T½-T½-T-T)" Then
                VGc = 2 : VGd = 3 : VGee = 3 : VGf = 2 : VGg = 2 : qtdeloop = 6
                I_1_0 = 0 : I_2_3 = 3 : I_4_9 = 9 : I_a5_13 = 13 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Raga Priyadharshini (T-T½-T½-T½-s)" Then
                VGc = 2 : VGd = 3 : VGee = 3 : VGf = 3 : VGg = 1 : qtdeloop = 6
                I_1_0 = 0 : I_2_3 = 3 : I_4_9 = 9 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Gorakh Kalyan (T-T½-2T-s-T)" Then
                VGc = 2 : VGd = 3 : VGee = 4 : VGf = 1 : VGg = 2 : qtdeloop = 6
                I_1_0 = 0 : I_2_3 = 3 : I_4_9 = 9 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Raga Rasranjani (T-T½-2T-T-s)" Then
                VGc = 2 : VGd = 3 : VGee = 4 : VGf = 2 : VGg = 1 : qtdeloop = 6
                I_1_0 = 0 : I_2_3 = 3 : I_4_9 = 9 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Bhavani tetratônica (T-T½-2T-T½)" Then
                VGc = 2 : VGd = 3 : VGee = 4 : VGf = 3 : qtdeloop = 5
                I_1_0 = 0 : I_2_3 = 3 : I_4_9 = 9 : I_6_15 = 15 : I_8_21 = 21
            ElseIf VGoo = "Raga Jaganmohanam (T-2T-s-s-T-T)" Then
                VGc = 2 : VGd = 4 : VGee = 1 : VGf = 1 : VGg = 2 : VGh = 2 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_a4_10 = 10 : I_5_12 = 12 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Raga Sarasvati (T-2T-s-T-s-T)" Then
                VGc = 2 : VGd = 4 : VGee = 1 : VGf = 3 : VGg = 1 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_a4_10 = 10 : I_5_12 = 12 : I_a6_16 = 16 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Shri Kalyan (T-2T-s-T-T½)" Then
                VGc = 2 : VGd = 4 : VGee = 1 : VGf = 2 : VGg = 3 : qtdeloop = 6
                I_1_0 = 0 : I_2_3 = 3 : I_a4_10 = 10 : I_5_12 = 12 : I_6_15 = 15 : I_8_21 = 21
            ElseIf VGoo = "Raga Sarasvati (T-2T-s-T½-s-s)" Then
                VGc = 2 : VGd = 4 : VGee = 1 : VGf = 3 : VGg = 1 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_2_3 = 3 : I_a4_10 = 10 : I_5_12 = 12 : I_a6_16 = 16 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Vaijayanti (T-2T-s-2T-s)" Then
                VGc = 2 : VGd = 4 : VGee = 1 : VGf = 4 : VGg = 1 : qtdeloop = 6
                I_1_0 = 0 : I_2_3 = 3 : I_a4_10 = 10 : I_5_12 = 12 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Sumukam (T-2T-2T½-s)" Then
                VGc = 2 : VGd = 4 : VGee = 5 : VGf = 1 : qtdeloop = 5
                I_1_0 = 0 : I_2_3 = 3 : I_a4_10 = 10 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Matha Kokila (T-2T½-T-s-T)" Then
                VGc = 2 : VGd = 5 : VGee = 2 : VGf = 1 : VGg = 2 : qtdeloop = 6
                I_1_0 = 0 : I_2_3 = 3 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Mela Yagapriya (31) (T½-s-s-T-s-s-T½)" Then
                VGc = 3 : VGd = 1 : VGee = 1 : VGf = 2 : VGg = 1 : VGh = 1 : VGi = 3 : qtdeloop = 8
                I_1_0 = 0 : I_a2_4 = 4 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_6_15 = 15 : I_8_21 = 21
            ElseIf VGoo = "Mela Gangeyabhusani (33) (T½-s-s-T-s-T½-s)" OrElse VGoo = "Sengah (T½-s-s-T-s-T½-s)" Then
                VGc = 3 : VGd = 1 : VGee = 1 : VGf = 2 : VGg = 1 : VGh = 3 : VGi = 1 : qtdeloop = 8
                If VGoo = "Sengah (T½-s-s-T-s-T½-s)" Then I_1_0 = 0 : I_b3_5 = 5 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
                If VGoo = "Mela Gangeyabhusani (33) (T½-s-s-T-s-T½-s)" Then I_1_0 = 0 : I_a2_4 = 4 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Mela Gangeyabhusani (33)", "Sengah"})
            ElseIf VGoo = "Mela Vagadhisvari (34) (T½-s-s-T-T-s-T)" Then
                VGc = 3 : VGd = 1 : VGee = 1 : VGf = 2 : VGg = 2 : VGh = 1 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_a2_4 = 4 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Houzam (T½-s-s-T-T-T-s)" OrElse VGoo = "Mela Sulini (35) (T½-s-s-T-T-T-s)" Then
                VGc = 3 : VGd = 1 : VGee = 1 : VGf = 2 : VGg = 2 : VGh = 2 : VGi = 1 : qtdeloop = 8
                If VGoo = "Mela Sulini (35) (T½-s-s-T-T-T-s)" Then I_1_0 = 0 : I_a2_4 = 4 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
                If VGoo = "Houzam (T½-s-s-T-T-T-s)" Then I_1_0 = 0 : I_b3_5 = 5 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Houzam", "Mela Sulini (35)"})
            ElseIf VGoo = "Mela Chalanata (36) (T½-s-s-T-T½-s-s)" Then
                VGc = 3 : VGd = 1 : VGee = 1 : VGf = 2 : VGg = 3 : VGh = 1 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_a2_4 = 4 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_a6_16 = 16 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Bhanumanjari (T½-s-s-T-T½-T)" Then
                VGc = 3 : VGd = 1 : VGee = 1 : VGf = 2 : VGg = 3 : VGh = 2 : qtdeloop = 7
                I_1_0 = 0 : I_a2_4 = 4 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Frigia cromática (T½-s-s-T½-s-s-T)" Then
                VGc = 3 : VGd = 1 : VGee = 1 : VGf = 3 : VGg = 1 : VGh = 1 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_a2_4 = 4 : I_3_6 = 6 : I_4_9 = 9 : I_a5_13 = 13 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Mela Sucharitra (67) (T½-s-T-s-s-s-T½)" Then
                VGc = 3 : VGd = 1 : VGee = 2 : VGf = 1 : VGg = 1 : VGh = 1 : VGi = 3 : qtdeloop = 8
                I_1_0 = 0 : I_a2_4 = 4 : I_3_6 = 6 : I_a4_10 = 10 : I_5_12 = 12 : I_b6_14 = 14 : I_6_15 = 15 : I_8_21 = 21
            ElseIf VGoo = "Mela Jyotisvarupini (68) (T½-s-T-s-s-T-T)" Then
                VGc = 3 : VGd = 1 : VGee = 2 : VGf = 1 : VGg = 1 : VGh = 2 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_a2_4 = 4 : I_3_6 = 6 : I_a4_10 = 10 : I_5_12 = 12 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Mela Dhatuvardhani (69) (T½-s-T-s-s-T½-s)" Then
                VGc = 3 : VGd = 1 : VGee = 2 : VGf = 1 : VGg = 1 : VGh = 3 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_a2_4 = 4 : I_3_6 = 6 : I_a4_10 = 10 : I_5_12 = 12 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Diminuída 2 (T½-s-T-s-T-s-T)" OrElse VGoo = "Mela Nasikabhusani (70) (T½-s-T-s-T-s-T)" OrElse VGoo = "Húngara Maior 1 (T½-s-T-s-T-s-T)" Then
                VGc = 3 : VGd = 1 : VGee = 2 : VGf = 1 : VGg = 2 : VGh = 1 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_a2_4 = 4 : I_3_6 = 6 : I_a4_10 = 10 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Diminuída 2", "Húngara Maior 1", "Mela Nasikabhusani (70)"})
            ElseIf VGoo = "Mela Kosalam (71) (T½-s-T-s-T-T-s)" Then
                VGc = 3 : VGd = 1 : VGee = 2 : VGf = 1 : VGg = 2 : VGh = 2 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_a2_4 = 4 : I_3_6 = 6 : I_a4_10 = 10 : I_5_12 = 12 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Mela Rasikapriya (72) (T½-s-T-s-T½-s-s)" OrElse VGoo = "Mela Visvambari (54) (T½-s-T-s-T½-s-s)" Then
                VGc = 3 : VGd = 1 : VGee = 2 : VGf = 1 : VGg = 3 : VGh = 1 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_a2_4 = 4 : I_3_6 = 6 : I_a4_10 = 10 : I_5_12 = 12 : I_a6_16 = 16 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Mela Rasikapriya (72)", "Mela Visvambari (54)"})
            ElseIf VGoo = "Raga Rasamanjari (T½-s-T-s-2T-s)" Then
                VGc = 3 : VGd = 1 : VGee = 2 : VGf = 1 : VGg = 4 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_a2_4 = 4 : I_3_6 = 6 : I_a4_10 = 10 : I_5_12 = 12 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Aumentada 2 (T½-s-T½-s-T-T)" Then
                VGc = 3 : VGd = 1 : VGee = 3 : VGf = 1 : VGg = 2 : VGh = 2 : qtdeloop = 7
                I_1_0 = 0 : I_a2_4 = 4 : I_3_6 = 6 : I_5_12 = 12 : I_a5_13 = 13 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Aumentada 1 (T½-s-T½-s-T½-s)" OrElse VGoo = "Genus tertium (T½-s-T½-s-T½-s)" Then
                VGc = 3 : VGd = 1 : VGee = 3 : VGf = 1 : VGg = 3 : VGh = 1 : qtdeloop = 7
                If VGoo = "Aumentada 1 (T½-s-T½-s-T½-s)" Then I_1_0 = 0 : I_b3_5 = 5 : I_3_6 = 6 : I_5_12 = 12 : I_a5_13 = 13 : I_7_18 = 18 : I_8_21 = 21
                If VGoo = "Genus tertium (T½-s-T½-s-T½-s)" Then I_1_0 = 0 : I_b3_5 = 5 : I_3_6 = 6 : I_5_12 = 12 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Aumentada 1", "Genus tertium"})
            ElseIf VGoo = "Raga Mohanangi (T½-s-T½-T-T½)" Then
                VGc = 3 : VGd = 1 : VGee = 3 : VGf = 2 : VGg = 3 : qtdeloop = 6
                I_1_0 = 0 : I_a2_4 = 4 : I_3_6 = 6 : I_5_12 = 12 : I_6_15 = 15 : I_8_21 = 21
            ElseIf VGoo = "Pentatônica Menor 3 (T½-T-s-2T-T)" OrElse VGoo = "Raga Jayakauns (T½-T-s-2T-T)" Then
                VGc = 3 : VGd = 2 : VGee = 1 : VGf = 4 : VGg = 2 : qtdeloop = 6
                I_1_0 = 0 : I_b3_5 = 5 : I_4_9 = 9 : I_b5_11 = 11 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Pentatônica Menor 3", "Raga Jayakauns"})
            ElseIf VGoo = "Frigia hexatônica (T½-T-T-s-T-T)" OrElse VGoo = "Raga Gopikavasantam (T½-T-T-s-T-T)" Then
                VGc = 3 : VGd = 2 : VGee = 2 : VGf = 1 : VGg = 2 : VGh = 2 : qtdeloop = 7
                I_1_0 = 0 : I_b3_5 = 5 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Frigia hexatônica", "Raga Gopikavasantam"})
            ElseIf VGoo = "Raga Takka (T½-T-T-s-T½-s)" Then
                VGc = 3 : VGd = 2 : VGee = 2 : VGf = 1 : VGg = 3 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_b3_5 = 5 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Kokil Pancham (T½-T-T-s-2T)" Then
                VGc = 3 : VGd = 2 : VGee = 2 : VGf = 1 : VGg = 4 : qtdeloop = 6
                I_1_0 = 0 : I_b3_5 = 5 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_8_21 = 21
            ElseIf VGoo = "Raga Manohari (T½-T-T-T-s-T)" Then
                VGc = 3 : VGd = 2 : VGee = 2 : VGf = 2 : VGg = 1 : VGh = 2 : qtdeloop = 7
                I_1_0 = 0 : I_b3_5 = 5 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Kyemyonjo (T½-T-T-T-T½)" Then
                VGc = 3 : VGd = 2 : VGee = 2 : VGf = 2 : VGg = 3 : qtdeloop = 6
                I_1_0 = 0 : I_b3_5 = 5 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_8_21 = 21
            ElseIf VGoo = "Raga Nata (T½-T-T-2T-s)" Then
                VGc = 3 : VGd = 2 : VGee = 2 : VGf = 4 : VGg = 1 : qtdeloop = 6
                I_1_0 = 0 : I_b3_5 = 5 : I_4_9 = 9 : I_5_12 = 12 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Malkauns (T½-T-T½-T-T)" OrElse VGoo = "Yi Ze (T½-T-T½-T-T)" Then
                VGc = 3 : VGd = 2 : VGee = 3 : VGf = 2 : VGg = 2 : qtdeloop = 6
                I_1_0 = 0 : I_b3_5 = 5 : I_4_9 = 9 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Raga Malkauns", "Yi Ze"})
            ElseIf VGoo = "Raga Chandrakauns-Kiravani (T½-T-T½-T½-s)" Then
                VGc = 3 : VGd = 2 : VGee = 3 : VGf = 3 : VGg = 1 : qtdeloop = 6
                I_1_0 = 0 : I_b3_5 = 5 : I_4_9 = 9 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Chandrakauns-Kafi (T½-T-2T-s-T)" Then
                VGc = 3 : VGd = 2 : VGee = 4 : VGf = 1 : VGg = 2 : qtdeloop = 6
                I_1_0 = 0 : I_b3_5 = 5 : I_4_9 = 9 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Raga Chandrakauns-Modern (T½-T-2T-T-s)" Then
                VGc = 3 : VGd = 2 : VGee = 4 : VGf = 2 : VGg = 1 : qtdeloop = 6
                I_1_0 = 0 : I_b3_5 = 5 : I_4_9 = 9 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Madhukauns (T½-T½-s-T-s-T)" Then
                VGc = 3 : VGd = 3 : VGee = 1 : VGf = 2 : VGg = 1 : VGh = 2 : qtdeloop = 7
                I_1_0 = 0 : I_b3_5 = 5 : I_a4_10 = 10 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Raga Jivantini (T½-T½-s-T½-s-s)" Then
                VGc = 3 : VGd = 3 : VGee = 1 : VGf = 3 : VGg = 1 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_b3_5 = 5 : I_a4_10 = 10 : I_5_12 = 12 : I_a6_16 = 16 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Samudhra Priya (T½-T½-s-T½-T)" Then
                VGc = 3 : VGd = 3 : VGee = 1 : VGf = 3 : VGg = 2 : qtdeloop = 6
                I_1_0 = 0 : I_b3_5 = 5 : I_a4_10 = 10 : I_5_12 = 12 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Raga Multani (T½-T½-s-2T-s)" Then
                VGc = 3 : VGd = 3 : VGee = 1 : VGf = 4 : VGg = 1 : qtdeloop = 6
                I_1_0 = 0 : I_b3_5 = 5 : I_a4_10 = 10 : I_5_12 = 12 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Chin (T½-T½-T-T-T)" OrElse VGoo = "Raga Harikauns (T½-T½-T-T-T)" Then
                VGc = 3 : VGd = 3 : VGee = 2 : VGf = 2 : VGg = 2 : qtdeloop = 6
                I_1_0 = 0 : I_b3_5 = 5 : I_b5_11 = 11 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Chin", "Raga Harikauns"})
            ElseIf VGoo = "Raga Shailaja (T½-2T-s-T-T)" Then
                VGc = 3 : VGd = 4 : VGee = 1 : VGf = 2 : VGg = 2 : qtdeloop = 6
                I_1_0 = 0 : I_b3_5 = 5 : I_5_12 = 12 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Bi Yu (T½-2T-T½-T)" Then
                VGc = 3 : VGd = 4 : VGee = 3 : VGf = 2 : qtdeloop = 5
                I_1_0 = 0 : I_b3_5 = 5 : I_5_12 = 12 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Peruana tritônica 2 (T½-2T-2T½)" Then
                VGc = 3 : VGd = 4 : VGee = 5 : qtdeloop = 4
                I_1_0 = 0 : I_b3_5 = 5 : I_5_12 = 12 : I_8_21 = 21
            ElseIf VGoo = "Ute (T½-3T½-T)" Then
                VGc = 3 : VGd = 7 : VGee = 2 : qtdeloop = 4
                I_1_0 = 0 : I_b3_5 = 5 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Raga Saravati (2T-s-T-s-s-T½)" Then
                VGc = 4 : VGd = 1 : VGee = 2 : VGf = 1 : VGg = 1 : VGh = 3 : qtdeloop = 7
                I_1_0 = 0 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_6_15 = 15 : I_8_21 = 21
            ElseIf VGoo = "Raga Kamalamanohari (2T-s-T-s-T-T)" Then
                VGc = 4 : VGd = 1 : VGee = 2 : VGf = 1 : VGg = 2 : VGh = 2 : qtdeloop = 7
                I_1_0 = 0 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Raga Paraju (2T-s-T-s-T½-s)" Then
                VGc = 4 : VGd = 1 : VGee = 2 : VGf = 1 : VGg = 3 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Zilaf (2T-s-T-s-2T)" Then
                VGc = 4 : VGd = 1 : VGee = 2 : VGf = 1 : VGg = 4 : qtdeloop = 6
                I_1_0 = 0 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_8_21 = 21
            ElseIf VGoo = "Raga Madhuri (2T-s-T-T-s-s-s)" Then
                VGc = 4 : VGd = 1 : VGee = 2 : VGf = 2 : VGg = 1 : VGh = 1 : VGi = 1 : qtdeloop = 8
                I_1_0 = 0 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Khamas (2T-s-T-T-s-T)" Then
                VGc = 4 : VGd = 1 : VGee = 2 : VGf = 2 : VGg = 1 : VGh = 2 : qtdeloop = 7
                I_1_0 = 0 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Genus secundum (2T-s-T-T-T-s)" OrElse VGoo = "Raga Hari Nata (2T-s-T-T-T-s)" Then
                VGc = 4 : VGd = 1 : VGee = 2 : VGf = 2 : VGg = 2 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Genus secundum", "Raga Hari Nata"})
            ElseIf VGoo = "Raga Nagasvaravali (2T-s-T-T-T½)" Then
                VGc = 4 : VGd = 1 : VGee = 2 : VGf = 2 : VGg = 3 : qtdeloop = 6
                I_1_0 = 0 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_8_21 = 21
            ElseIf VGoo = "Raga Tilang (2T-s-T-T½-s-s)" Then
                VGc = 4 : VGd = 1 : VGee = 2 : VGf = 3 : VGg = 1 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_a6_16 = 16 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Pentatônica de Dominante (2T-s-T-T½-T)" Then
                VGc = 4 : VGd = 1 : VGee = 2 : VGf = 3 : VGg = 2 : qtdeloop = 6
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_5_12 = 12 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Raga Gambhiranata (2T-s-T-2T-s)" OrElse VGoo = "Ryukyu (2T-s-T-2T-s)" Then
                VGc = 4 : VGd = 1 : VGee = 2 : VGf = 4 : VGg = 1 : qtdeloop = 6
                I_1_0 = 0 : I_3_6 = 6 : I_4_9 = 9 : I_5_12 = 12 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Raga Gambhiranata", "Ryukyu"})
            ElseIf VGoo = "Raga Girija (2T-s-T½-T½-s)" Then
                VGc = 4 : VGd = 1 : VGee = 3 : VGf = 3 : VGg = 1 : qtdeloop = 6
                I_1_0 = 0 : I_3_6 = 6 : I_4_9 = 9 : I_a5_13 = 13 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Khamaji Durga (2T-s-2T-s-T)" Then
                VGc = 4 : VGd = 1 : VGee = 4 : VGf = 1 : VGg = 2 : qtdeloop = 6
                I_1_0 = 0 : I_3_6 = 6 : I_4_9 = 9 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Kumoi (T-s-2T-T-T½)" OrElse VGoo = "Raga Bhinna Shadja (2T-s-2T-T-s)" Then
                VGc = 4 : VGd = 1 : VGee = 4 : VGf = 2 : VGg = 1 : qtdeloop = 6
                I_1_0 = 0 : I_3_6 = 6 : I_4_9 = 9 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Kumoi", "Raga Bhinna Shadja"})
            ElseIf VGoo = "Raga Jyoti (2T-T-s-s-T-T)" Then
                VGc = 4 : VGd = 2 : VGee = 1 : VGf = 1 : VGg = 2 : VGh = 2 : qtdeloop = 7
                I_1_0 = 0 : I_3_6 = 6 : I_a4_10 = 10 : I_5_12 = 12 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Raga Vutari (2T-T-s-T-s-T)" Then
                VGc = 4 : VGd = 2 : VGee = 1 : VGf = 2 : VGg = 1 : VGh = 2 : qtdeloop = 7
                I_1_0 = 0 : I_3_6 = 6 : I_a4_10 = 10 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Raga Dhavalashri (2T-T-s-T-T½)" Then
                VGc = 4 : VGd = 2 : VGee = 1 : VGf = 2 : VGg = 3 : qtdeloop = 6
                I_1_0 = 0 : I_3_6 = 6 : I_a4_10 = 10 : I_5_12 = 12 : I_6_15 = 15 : I_8_21 = 21
            ElseIf VGoo = "Raga Vijayavasanta (2T-T-s-T½-s-s)" Then
                VGc = 4 : VGd = 2 : VGee = 1 : VGf = 3 : VGg = 1 : VGh = 1 : qtdeloop = 7
                I_1_0 = 0 : I_3_6 = 6 : I_a4_10 = 10 : I_5_12 = 12 : I_a6_16 = 16 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Amritavarsini (2T-T-s-2T-s)" OrElse VGoo = "China 1 (2T-T-s-2T-s)" Then
                VGc = 4 : VGd = 2 : VGee = 1 : VGf = 4 : VGg = 1 : qtdeloop = 6
                I_1_0 = 0 : I_3_6 = 6 : I_a4_10 = 10 : I_5_12 = 12 : I_7_18 = 18 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Raga Amritavarsini", "China 1"})
            ElseIf VGoo = "Raga Hindol (2T-T-T½-T-s)" Then
                VGc = 4 : VGd = 2 : VGee = 3 : VGf = 2 : VGg = 1 : qtdeloop = 6
                I_1_0 = 0 : I_3_6 = 6 : I_b5_11 = 11 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Valaji (2T-T½-T-s-T)" Then
                VGc = 4 : VGd = 3 : VGee = 2 : VGf = 1 : VGg = 2 : qtdeloop = 6
                I_1_0 = 0 : I_3_6 = 6 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Raga Mamata (2T-T½-T-T-s)" Then
                VGc = 4 : VGd = 3 : VGee = 2 : VGf = 2 : VGg = 1 : qtdeloop = 6
                I_1_0 = 0 : I_3_6 = 6 : I_5_12 = 12 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Mahathi (2T-T½-T½-T)" Then
                VGc = 4 : VGd = 3 : VGee = 3 : VGf = 2 : qtdeloop = 5
                I_1_0 = 0 : I_3_6 = 6 : I_5_12 = 12 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Peruana tritônica 1 (2T-T½-2T½)" OrElse VGoo = "Raga Malasri (2T-T½-2T½)" Then
                VGc = 4 : VGd = 3 : VGee = 5 : qtdeloop = 4
                I_1_0 = 0 : I_3_6 = 6 : I_5_12 = 12 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Peruana tritônica 1", "Raga Malasri"})
            ElseIf VGoo = "Tetratônica (2T½-s-2T½-s)" Then
                VGc = 5 : VGd = 1 : VGee = 5 : VGf = 1 : qtdeloop = 5
                I_1_0 = 0 : I_4_9 = 9 : I_b5_11 = 11 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Devranjani (2T½-T-s-T-T)" Then
                VGc = 5 : VGd = 2 : VGee = 1 : VGf = 2 : VGg = 2 : qtdeloop = 6
                I_1_0 = 0 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Raga Devaranji (2T½-T-s-T½-s)" Then
                VGc = 5 : VGd = 2 : VGee = 1 : VGf = 3 : VGg = 1 : qtdeloop = 6
                I_1_0 = 0 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Kuntvarali (2T½-T-T-s-T)" Then
                VGc = 5 : VGd = 2 : VGee = 2 : VGf = 1 : VGg = 2 : qtdeloop = 6
                I_1_0 = 0 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Raga Puruhutika (2T½-T-T-T-s)" Then
                VGc = 5 : VGd = 2 : VGee = 2 : VGf = 2 : VGg = 1 : qtdeloop = 6
                I_1_0 = 0 : I_4_9 = 9 : I_5_12 = 12 : I_6_15 = 15 : I_7_18 = 18 : I_8_21 = 21
            ElseIf VGoo = "Raga Sarvasri (2T½-T-2T½)" OrElse VGoo = "Tritônica (2T½-T-2T½)" OrElse VGoo = "Warao tritônica (2T½-T-2T½)" Then
                VGc = 5 : VGd = 2 : VGee = 5 : qtdeloop = 4
                I_1_0 = 0 : I_4_9 = 9 : I_5_12 = 12 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Raga Sarvasri", "Tritônica", "Warao tritônica"})
            ElseIf VGoo = "Sansagari (2T½-2T½-T)" Then
                VGc = 5 : VGd = 5 : VGee = 2 : qtdeloop = 4
                I_1_0 = 0 : I_4_9 = 9 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Honchoshi (2T½-3T½)" Then
                VGc = 5 : VGd = 7 : qtdeloop = 3
                I_1_0 = 0 : I_4_9 = 9 : I_8_21 = 21
            ElseIf VGoo = "Raga Ongkari (3T-s-2T½)" Then
                VGc = 6 : VGd = 1 : VGee = 5 : qtdeloop = 4
                I_1_0 = 0 : I_a4_10 = 10 : I_5_12 = 12 : I_8_21 = 21
            ElseIf VGoo = "Niagari ditônica (3T½-2T½)" Then
                VGc = 7 : VGd = 5 : qtdeloop = 3
                I_1_0 = 0 : I_5_12 = 12 : I_8_21 = 21
            ElseIf VGoo = "Mela Jhankaradhvani (19) (T-s-T-T-s-s-T½)" Then
                VGc = 2 : VGd = 1 : VGee = 2 : VGf = 2 : VGg = 1 : VGh = 1 : VGi = 3 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_b3_5 = 5 : I_4_9 = 9 : I_5_12 = 12 : I_b6_14 = 14 : I_6_15 = 15 : I_8_21 = 21
            ElseIf VGoo = "Mela Jalarnavam (38) (s-s-2T-s-s-T-T)" Then
                VGc = 1 : VGd = 1 : VGee = 4 : VGf = 1 : VGg = 1 : VGh = 2 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_b2_2 = 2 : I_2_3 = 3 : I_a4_10 = 10 : I_5_12 = 12 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
            ElseIf VGoo = "Mela Risabhapriya (62) (T-T-T-s-s-T-T)" OrElse VGoo = "Lidia Menor (T-T-T-s-s-T-T)" Then
                VGc = 2 : VGd = 2 : VGee = 2 : VGf = 1 : VGg = 1 : VGh = 2 : VGi = 2 : qtdeloop = 8
                I_1_0 = 0 : I_2_3 = 3 : I_3_6 = 6 : I_a4_10 = 10 : I_5_12 = 12 : I_b6_14 = 14 : I_b7_17 = 17 : I_8_21 = 21
                Teclado.ListBox1.Items.AddRange(New Object() {"Mela Risabhapriya (62)", "Lídio Menor"})
            End If

            Intervalos = "Intervalos: "
            Intervalos2 = "("
            If I_1_0 = 0 Then Intervalos = Intervalos & "1," : Intervalos2 = Intervalos2 & "U,"
            If I_a1_1 = 1 Then Intervalos = Intervalos & "#1," : Intervalos2 = Intervalos2 & "1a,"
            If I_b2_2 = 2 Then Intervalos = Intervalos & "b2," : Intervalos2 = Intervalos2 & "2m,"
            If I_2_3 = 3 Then Intervalos = Intervalos & "2," : Intervalos2 = Intervalos2 & "2M,"
            If I_a2_4 = 4 Then Intervalos = Intervalos & "#2," : Intervalos2 = Intervalos2 & "2a,"
            If I_b3_5 = 5 Then Intervalos = Intervalos & "b3," : Intervalos2 = Intervalos2 & "3m,"
            If I_3_6 = 6 Then Intervalos = Intervalos & "3," : Intervalos2 = Intervalos2 & "3M,"
            If I_a3_7 = 7 Then Intervalos = Intervalos & "#3," : Intervalos2 = Intervalos2 & "3a,"
            If I_b4_8 = 8 Then Intervalos = Intervalos & "b4," : Intervalos2 = Intervalos2 & "4d,"
            If I_4_9 = 9 Then Intervalos = Intervalos & "4," : Intervalos2 = Intervalos2 & "4J,"
            If I_a4_10 = 10 Then Intervalos = Intervalos & "#4," : Intervalos2 = Intervalos2 & "4a,"
            If I_b5_11 = 11 Then Intervalos = Intervalos & "b5," : Intervalos2 = Intervalos2 & "5d,"
            If I_5_12 = 12 Then Intervalos = Intervalos & "5," : Intervalos2 = Intervalos2 & "5J,"
            If I_a5_13 = 13 Then Intervalos = Intervalos & "#5," : Intervalos2 = Intervalos2 & "5a,"
            If I_b6_14 = 14 Then Intervalos = Intervalos & "b6," : Intervalos2 = Intervalos2 & "6m,"
            If I_6_15 = 15 Then Intervalos = Intervalos & "6," : Intervalos2 = Intervalos2 & "6M,"
            If I_a6_16 = 16 Then Intervalos = Intervalos & "#6," : Intervalos2 = Intervalos2 & "6a,"
            If I_b7_17 = 17 Then Intervalos = Intervalos & "b7," : Intervalos2 = Intervalos2 & "7m,"
            If I_7_18 = 18 Then Intervalos = Intervalos & "7," : Intervalos2 = Intervalos2 & "7M,"
            If I_a7_19 = 19 Then Intervalos = Intervalos & "#7," : Intervalos2 = Intervalos2 & "7a,"
            If I_b8_20 = 20 Then Intervalos = Intervalos & "b8," : Intervalos2 = Intervalos2 & "8d,"
            Intervalos = Intervalos & "8    " & Intervalos2 & "8J)"


            Teclado.ListBox1.Sorted = CBool(1) 'classifica em ordem crescente
            If Teclado.ListBox1.Items.Count = 0 Then Teclado.ListBox1.Items.Add("Nenhuma")
            If VGoo <> "" Then exibeControles()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub PictureBox20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox20.Click, PictureBox1.Click

        Try

            If VGoo <> "" Then
                PictureBox1.Visible = Not PictureBox1.Visible
                RichTextBox1.Text = ""
                RichTextBox1.Visible = Not RichTextBox1.Visible
                PictureBox1.Top = 75
                RichTextBox1.Top = PictureBox1.Top + 50

                If VGoo = "Jônio" OrElse VGoo = "Dórico" OrElse VGoo = "Frígio" OrElse VGoo = "Lídio" OrElse VGoo = "Mixolídio" OrElse VGoo = "Eólio" OrElse VGoo = "Lócrio" Then
                    RichTextBox1.LoadFile("Conteudo\Escalas\Escalas Gregas.rtf")
                End If
                'RichTextBox1.LoadFile("Conteudo\Escalas\Escala Maior.rtf")
                PictureBox1.BringToFront()
                RichTextBox1.BringToFront()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub PictureBox4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox4.DoubleClick, PictureBox3.DoubleClick

        Try

            ColorDialog1.ShowDialog()
            My.Settings.NovaCor = ColorDialog1.Color
            exibeControles()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub CorC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CorC.DoubleClick

        Try

            ColorDialogC.ShowDialog()
            My.Settings.NovaCorC = ColorDialogC.Color
            CorC.BackColor = ColorDialogC.Color
            exibeControles()
            Teclado.atualizaTeclado()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub CorD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CorD.DoubleClick

        Try

            ColorDialogD.ShowDialog()
            My.Settings.NovaCorD = ColorDialogD.Color
            CorD.BackColor = ColorDialogD.Color
            exibeControles()
            Teclado.atualizaTeclado()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub CorE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CorE.DoubleClick

        Try

            ColorDialogE.ShowDialog()
            My.Settings.NovaCorE = ColorDialogE.Color
            CorE.BackColor = ColorDialogE.Color
            exibeControles()
            Teclado.atualizaTeclado()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub CorF_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CorF.DoubleClick

        Try

            ColorDialogF.ShowDialog()
            My.Settings.NovaCorF = ColorDialogF.Color
            CorF.BackColor = ColorDialogF.Color
            exibeControles()
            Teclado.atualizaTeclado()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub CorG_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CorG.DoubleClick

        Try

            ColorDialogG.ShowDialog()
            My.Settings.NovaCorG = ColorDialogG.Color
            CorG.BackColor = ColorDialogG.Color
            exibeControles()
            Teclado.atualizaTeclado()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub CorA_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CorA.DoubleClick

        Try

            ColorDialogA.ShowDialog()
            My.Settings.NovaCorA = ColorDialogA.Color
            CorA.BackColor = ColorDialogA.Color
            exibeControles()
            Teclado.atualizaTeclado()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub CorB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CorB.DoubleClick

        Try

            ColorDialogB.ShowDialog()
            My.Settings.NovaCorB = ColorDialogB.Color
            CorB.BackColor = ColorDialogB.Color
            exibeControles()
            Teclado.atualizaTeclado()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub PictureBox6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox6.DoubleClick

        Try

            ColorDialogC.Color = Color.FromArgb(2, 0, 0)
            ColorDialogD.Color = Color.FromArgb(51, 153, 102)
            ColorDialogE.Color = Color.FromArgb(0, 153, 255)
            ColorDialogF.Color = Color.FromArgb(255, 0, 0)
            ColorDialogG.Color = Color.FromArgb(217, 214, 89)
            ColorDialogA.Color = Color.FromArgb(255, 153, 0)
            ColorDialogB.Color = Color.FromArgb(255, 153, 255)

            exibeControles()
            Teclado.atualizaTeclado()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub PictureBox7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox7.Click

        Try

            For Each Item As ToolStripMenuItem In Teclado.MenuStrip1.Items
                If TypeOf (Item) Is ToolStripMenuItem Then 'ignora tudo o que não for ToolStripMenuItem
                    If Item.Text <> "Escalas" Then
                        Escolha_uma_escala.ListBox2.Items.Add(Item.Text)
                    End If
                    preenchelistbox(Item)
                End If
            Next
            Escolha_uma_escala.ListBox2.Sorted = CBool(1) 'classifica em ordem crescente
            Escolha_uma_escala.Show()
            Escolha_uma_escala.BringToFront()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Public Sub preenchelistbox(ByVal ToolstripMenuItem As ToolStripMenuItem)

        Try

            'Faz a mesma coisa que a rotina do load
            For ii = 0 To ToolstripMenuItem.DropDownItems.Count - 1
                If TypeOf (ToolstripMenuItem.DropDownItems.Item(ii)) Is ToolStripMenuItem Then 'ignora tudo o que não for ToolStripMenuItem
                    If ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Exibir notas com cores diferentes" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Inverter enarmônia" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Usar cifras simples. Ex.: F## = G, Cb = B, E# = F etc" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Tocar escalas 1 oitava acima" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Tocar escalas 1 oitava abaixo" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Minimizar" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Fechar" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Blues" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Exóticas A-E" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Árabes" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Aumentadas" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Balinesas" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Be-Bop" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Chinesas" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Coreanas" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Diminuídas (Diminished)" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Enigmáticas" AndAlso _
                        ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Espanholas" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Esquimais" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Etíopes" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Exóticas F-J" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Genus" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Ghanesas" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Hawayanas" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Húngaras" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Israelitas" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Japonesas" AndAlso _
                        ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Javanesas" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Judaicas" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Exóticas K-O" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Maqam" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Mela 1-20" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Mela 21-40" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Mela 41-60" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Mela 61-72" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Messiânicas" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Mischung" AndAlso _
                        ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Napolitanas" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Niagari" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Oriental" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Exóticas P-T" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Persas" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Peruanas" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Raga A-B" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Raga C-D" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Raga E-H" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Raga I-L" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Raga M-N" AndAlso _
                        ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Raga O-R" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Raga S-T" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Raga U-Z" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Simétricas" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Skriabin" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Theta" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Todi" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Exóticas U-Z" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Warao" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Yu" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Gregas" AndAlso _
                        ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Prometheus" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Bizantinas" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Gregas Alteradas" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Harmônicas" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Jazz" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Menores" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Octatônicas" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Pentatônicas" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Tons Inteiros" AndAlso ToolstripMenuItem.DropDownItems.Item(ii).Text <> "Nenhuma" Then
                        Escolha_uma_escala.ListBox2.Items.Add(ToolstripMenuItem.DropDownItems.Item(ii).Text)
                    End If
                    'Aqui é a parte mais importante. Isso se chama recursão, onde um método
                    'chama ele mesmo, e passa o próprio item que está sendo adicionado ao listbox
                    'Assim ele adiciona o item e chama a mesma rotina para adicionar os seus respectivos
                    'subItems
                    preenchelistbox(CType(ToolstripMenuItem.DropDownItems.Item(ii), Windows.Forms.ToolStripMenuItem))
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub TocaEscalas(ByVal sender As Object, ByVal e As System.EventArgs) Handles EscalaA1.Click, EscalaA2.Click, EscalaB1.Click, EscalaB2.Click, EscalaC1.Click, EscalaC2.Click, EscalaD1.Click, EscalaD2.Click, EscalaE1.Click, EscalaE2.Click, EscalaF1.Click, EscalaF2.Click, EscalaG1.Click, EscalaG2.Click, EscalaH1.Click, EscalaH2.Click, EscalaI1.Click, EscalaI2.Click, EscalaJ1.Click, EscalaJ2.Click, EscalaK1.Click, EscalaK2.Click, EscalaL1.Click, EscalaL2.Click

        Try

            If VGoo <> "" Then

                ' Obtém referência ao picturebox que invocou este método.
                Dim pbox As PictureBox = DirectCast(sender, PictureBox)
                If pbox.Name = "EscalaA1" OrElse pbox.Name = "EscalaA2" Then

                    If pbox.Name = "EscalaA1" Then VGaa = 1 : VGz = 1 : VGx = 1
                    If pbox.Name = "EscalaA2" Then VGaa = 2 : VGz = qtdeloop - 1 : VGx = 1 + VGc + VGd + VGee + VGf + VGg + VGh + VGi + VGj + VGk + VGl + VGm + VGn + VGt
                ElseIf pbox.Name = "EscalaF1" OrElse pbox.Name = "EscalaF2" Then

                    If pbox.Name = "EscalaF1" Then VGaa = 1 : VGz = 1 : VGx = 6
                    If pbox.Name = "EscalaF2" Then VGaa = 2 : VGz = qtdeloop - 1 : VGx = 6 + VGc + VGd + VGee + VGf + VGg + VGh + VGi + VGj + VGk + VGl + VGm + VGn + VGt
                ElseIf pbox.Name = "EscalaK1" OrElse pbox.Name = "EscalaK2" Then

                    If pbox.Name = "EscalaK1" Then VGaa = 1 : VGz = 1 : VGx = 11
                    If pbox.Name = "EscalaK2" Then VGaa = 2 : VGz = qtdeloop - 1 : VGx = 11 + VGc + VGd + VGee + VGf + VGg + VGh + VGi + VGj + VGk + VGl + VGm + VGn + VGt
                ElseIf pbox.Name = "EscalaD1" OrElse pbox.Name = "EscalaD2" Then

                    If pbox.Name = "EscalaD1" Then VGaa = 1 : VGz = 1 : VGx = 4
                    If pbox.Name = "EscalaD2" Then VGaa = 2 : VGz = qtdeloop - 1 : VGx = 4 + VGc + VGd + VGee + VGf + VGg + VGh + VGi + VGj + VGk + VGl + VGm + VGn + VGt
                ElseIf pbox.Name = "EscalaI1" OrElse pbox.Name = "EscalaI2" Then

                    If pbox.Name = "EscalaI1" Then VGaa = 1 : VGz = 1 : VGx = 9
                    If pbox.Name = "EscalaI2" Then VGaa = 2 : VGz = qtdeloop - 1 : VGx = 9 + VGc + VGd + VGee + VGf + VGg + VGh + VGi + VGj + VGk + VGl + VGm + VGn + VGt
                ElseIf pbox.Name = "EscalaB1" OrElse pbox.Name = "EscalaB2" Then

                    If pbox.Name = "EscalaB1" Then VGaa = 1 : VGz = 1 : VGx = 2
                    If pbox.Name = "EscalaB2" Then VGaa = 2 : VGz = qtdeloop - 1 : VGx = 2 + VGc + VGd + VGee + VGf + VGg + VGh + VGi + VGj + VGk + VGl + VGm + VGn + VGt
                ElseIf pbox.Name = "EscalaG1" OrElse pbox.Name = "EscalaG2" Then

                    If pbox.Name = "EscalaG1" Then VGaa = 1 : VGz = 1 : VGx = 7
                    If pbox.Name = "EscalaG2" Then VGaa = 2 : VGz = qtdeloop - 1 : VGx = 7 + VGc + VGd + VGee + VGf + VGg + VGh + VGi + VGj + VGk + VGl + VGm + VGn + VGt
                ElseIf pbox.Name = "EscalaL1" OrElse pbox.Name = "EscalaL2" Then

                    If pbox.Name = "EscalaL1" Then VGaa = 1 : VGz = 1 : VGx = 12
                    If pbox.Name = "EscalaL2" Then VGaa = 2 : VGz = qtdeloop - 1 : VGx = 12 + VGc + VGd + VGee + VGf + VGg + VGh + VGi + VGj + VGk + VGl + VGm + VGn + VGt
                ElseIf pbox.Name = "EscalaE1" OrElse pbox.Name = "EscalaE2" Then

                    If pbox.Name = "EscalaE1" Then VGaa = 1 : VGz = 1 : VGx = 5
                    If pbox.Name = "EscalaE2" Then VGaa = 2 : VGz = qtdeloop - 1 : VGx = 5 + VGc + VGd + VGee + VGf + VGg + VGh + VGi + VGj + VGk + VGl + VGm + VGn + VGt
                ElseIf pbox.Name = "EscalaJ1" OrElse pbox.Name = "EscalaJ2" Then

                    If pbox.Name = "EscalaJ1" Then VGaa = 1 : VGz = 1 : VGx = 10
                    If pbox.Name = "EscalaJ2" Then VGaa = 2 : VGz = qtdeloop - 1 : VGx = 10 + VGc + VGd + VGee + VGf + VGg + VGh + VGi + VGj + VGk + VGl + VGm + VGn + VGt
                ElseIf pbox.Name = "EscalaC1" OrElse pbox.Name = "EscalaC2" Then

                    If pbox.Name = "EscalaC1" Then VGaa = 1 : VGz = 1 : VGx = 3
                    If pbox.Name = "EscalaC2" Then VGaa = 2 : VGz = qtdeloop - 1 : VGx = 3 + VGc + VGd + VGee + VGf + VGg + VGh + VGi + VGj + VGk + VGl + VGm + VGn + VGt
                ElseIf pbox.Name = "EscalaH1" OrElse pbox.Name = "EscalaH2" Then

                    If pbox.Name = "EscalaH1" Then VGaa = 1 : VGz = 1 : VGx = 8
                    If pbox.Name = "EscalaH2" Then VGaa = 2 : VGz = qtdeloop - 1 : VGx = 8 + VGc + VGd + VGee + VGf + VGg + VGh + VGi + VGj + VGk + VGl + VGm + VGn + VGt
                End If
                Teclado.atualizaTeclado()
                Teclado.Timer1.Interval = 1
                Teclado.Timer1.Enabled = True
                VGzz = 0
                Teclado.Timer2.Enabled = True
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub TocaAcordes(ByVal sender As Object, ByVal e As System.EventArgs) Handles TocaAcordesEscalaDó.Click, TocaAcordesEscalaSol.Click, TocaAcordesEscalaRé.Click, TocaAcordesEscalaLá.Click, TocaAcordesEscalaMi.Click, TocaAcordesEscalaSi.Click, TocaAcordesEscalaFáSus.Click, TocaAcordesEscalaDóSus.Click, TocaAcordesEscalaSolSus.Click, TocaAcordesEscalaRéSus.Click, TocaAcordesEscalaLáSus.Click, TocaAcordesEscalaFá.Click

        Try

            VGhhh = 1

            If VGoo <> "" Then
                a_I = 0
                a_II = 0
                a_III = 0
                a_IV = 0
                a_V = 0
                a_VI = 0
                a_VII = 0
                VGab = 1

                ' Obtém referência ao picturebox que invocou este método.
                Dim pbox As PictureBox = DirectCast(sender, PictureBox)
                If pbox.Name = "TocaAcordesEscalaDó" Then
                    VGjjj = "Dó"
                    If VGccc = "Progressão 1" Then
                        Teclado.Prog1A.Text = "C"
                        Teclado.Prog2A.Text = "Dm"
                        Teclado.Prog3A.Text = "Em"
                        Teclado.Prog4A.Text = "F"
                        Teclado.Prog5A.Text = "G"
                        Teclado.Prog6A.Text = "Am"
                        Teclado.Prog7A.Text = "Bº"
                    End If
                ElseIf pbox.Name = "TocaAcordesEscalaSol" Then
                    VGjjj = "Sol"
                    If VGccc = "Progressão 1" Then
                        Teclado.Prog1A.Text = "G"
                        Teclado.Prog2A.Text = "Am"
                        Teclado.Prog3A.Text = "Bm"
                        Teclado.Prog4A.Text = "C"
                        Teclado.Prog5A.Text = "D"
                        Teclado.Prog6A.Text = "Em"
                        Teclado.Prog7A.Text = "F#º"
                    End If
                ElseIf pbox.Name = "TocaAcordesEscalaRé" Then
                    VGjjj = "Ré"
                    If VGccc = "Progressão 1" Then
                        Teclado.Prog1A.Text = "D"
                        Teclado.Prog2A.Text = "Em"
                        Teclado.Prog3A.Text = "F#m"
                        Teclado.Prog4A.Text = "G"
                        Teclado.Prog5A.Text = "A"
                        Teclado.Prog6A.Text = "Bm"
                        Teclado.Prog7A.Text = "C#º"
                    End If
                ElseIf pbox.Name = "TocaAcordesEscalaLá" Then
                    VGjjj = "Lá"
                ElseIf pbox.Name = "TocaAcordesEscalaMi" Then
                    VGjjj = "Mi"
                ElseIf pbox.Name = "TocaAcordesEscalaSi" Then
                    VGjjj = "Si"
                ElseIf pbox.Name = "TocaAcordesEscalaFáSus" Then
                    If Teclado.CheckBox200.Checked Then
                        VGjjj = "SolBemol"
                    Else
                        VGjjj = "FáSustenido"
                    End If
                ElseIf pbox.Name = "TocaAcordesEscalaDóSus" Then
                    If Teclado.CheckBox200.Checked Then
                        VGjjj = "RéBemol"
                    Else
                        VGjjj = "DóSustenido"
                    End If
                ElseIf pbox.Name = "TocaAcordesEscalaSolSus" Then
                    If Teclado.CheckBox200.Checked Then
                        VGjjj = "LáBemol"
                    Else
                        VGjjj = "SolSustenido"
                    End If
                ElseIf pbox.Name = "TocaAcordesEscalaRéSus" Then
                    If Teclado.CheckBox200.Checked Then
                        VGjjj = "MiBemol"
                    Else
                        VGjjj = "RéSustenido"
                    End If
                ElseIf pbox.Name = "TocaAcordesEscalaLáSus" Then
                    If Teclado.CheckBox200.Checked Then
                        VGjjj = "SiBemol"
                    Else
                        VGjjj = "LáSustenido"
                    End If
                ElseIf pbox.Name = "TocaAcordesEscalaFá" Then
                    VGjjj = "Fá"
                End If
                Call Teclado.atualizaTeclado()
                Teclado.Timer4.Interval = 1

                If Teclado.PictureBox5.Image Is Nothing Then
                Else
                    Teclado.Timer4.Enabled = True
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub Escalas_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        Teclado.WindowState = Me.WindowState
    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click
        Me.WindowState = CType(1, FormWindowState) '1 é para minimizar
        Teclado.WindowState = CType(1, FormWindowState) '1 é para minimizar
    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click

        Try

            Me.Close()
            Teclado.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub
End Class