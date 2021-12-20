Option Strict On
Option Explicit On
Imports System.Xml

Public Class TreinamentoDinâmicaMãoEsquerdaDireita
    Inherits PerPixelAlphaForm 'para mecher no form no modo design é preciso mudar para 'Inherits Form'
    'e depois de ajustado voltar para 'Inherits PerPixelAlphaForm' e clicar para mudar o inherit
    Public TransAmount As Byte = 255

    Dim ArrayTeclas(5, 88), ValorTopo(1), AjusteTeclasPretas, i As Integer
    Dim CorKeyVelocity(1, 88) As SolidBrush

    Dim Fonte As New Font("Arial", 9, FontStyle.Regular, GraphicsUnit.Pixel)
    Dim CorFonte As SolidBrush = New SolidBrush(Color.FromArgb(2, 0, 0))
    Dim CorFonte2 As SolidBrush = New SolidBrush(Color.FromArgb(250, 255, 255))

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

    Private Sub Desenhar()

        Try

            DefineCor()

            Dim gr As Graphics
            Dim FaceBit As New Bitmap(My.Resources.TreinamentoDinamica)
            gr = Graphics.FromImage(FaceBit)
            gr.SmoothingMode = SmoothingMode.AntiAlias

            Dim VPPP As SolidBrush = New SolidBrush(PPP.Color)
            Dim VPP As SolidBrush = New SolidBrush(PP.Color)
            Dim VP As SolidBrush = New SolidBrush(P.Color)
            Dim VMP As SolidBrush = New SolidBrush(MP.Color)
            Dim VMF As SolidBrush = New SolidBrush(MF.Color)
            Dim VF As SolidBrush = New SolidBrush(F.Color)
            Dim VFF As SolidBrush = New SolidBrush(FF.Color)
            Dim VFFF As SolidBrush = New SolidBrush(FFF.Color)

            Dim MedidasTexto As SizeF = gr.MeasureString(CStr(ArrayTeclas(0, i)), Fonte)
            Dim MedidasTexto2 As SizeF = gr.MeasureString(CStr(ArrayTeclas(3, i)), Fonte)
            Dim PosicaoX As Integer = CInt(((Me.Width) / 2) - (MedidasTexto.Width / 2))
            Dim PosicaoX2 As Integer = CInt(((Me.Width) / 2) - (MedidasTexto2.Width / 2))
            ArrayTeclas(4, i) = PosicaoX
            ArrayTeclas(5, i) = PosicaoX2

            If ArrayTeclas(0, 1) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 1), 93, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 1)), Fonte, CorFonte, ArrayTeclas(4, 1) - 434, ValorTopo(0) - 12)
            If ArrayTeclas(1, 1) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 1), 93, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 1)), Fonte, CorFonte, ArrayTeclas(5, 1) - 434, ValorTopo(1) - 12)
            If ArrayTeclas(0, 2) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 2), 101, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 2)), Fonte, CorFonte2, ArrayTeclas(4, 2) - 426, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 2) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 2), 101, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 2)), Fonte, CorFonte2, ArrayTeclas(5, 2) - 426, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 3) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 3), 110, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 3)), Fonte, CorFonte, ArrayTeclas(4, 3) - 417, ValorTopo(0) - 12)
            If ArrayTeclas(1, 3) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 3), 110, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 3)), Fonte, CorFonte, ArrayTeclas(5, 3) - 417, ValorTopo(1) - 12)
            If ArrayTeclas(0, 4) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 4), 127, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 4)), Fonte, CorFonte, ArrayTeclas(4, 4) - 400, ValorTopo(0) - 12)
            If ArrayTeclas(1, 4) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 4), 127, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 4)), Fonte, CorFonte, ArrayTeclas(5, 4) - 400, ValorTopo(1) - 12)
            If ArrayTeclas(0, 5) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 5), 135, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 5)), Fonte, CorFonte2, ArrayTeclas(4, 5) - 392, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 5) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 5), 135, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 5)), Fonte, CorFonte2, ArrayTeclas(5, 5) - 392, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 6) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 6), 144, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 6)), Fonte, CorFonte, ArrayTeclas(4, 6) - 383, ValorTopo(0) - 12)
            If ArrayTeclas(1, 6) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 6), 144, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 6)), Fonte, CorFonte, ArrayTeclas(5, 6) - 383, ValorTopo(1) - 12)
            If ArrayTeclas(0, 7) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 7), 152, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 7)), Fonte, CorFonte2, ArrayTeclas(4, 7) - 375, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 7) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 7), 152, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 7)), Fonte, CorFonte2, ArrayTeclas(5, 7) - 375, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 8) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 8), 161, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 8)), Fonte, CorFonte, ArrayTeclas(4, 8) - 366, ValorTopo(0) - 12)
            If ArrayTeclas(1, 8) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 8), 161, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 8)), Fonte, CorFonte, ArrayTeclas(5, 8) - 366, ValorTopo(1) - 12)
            If ArrayTeclas(0, 9) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 9), 178, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 9)), Fonte, CorFonte, ArrayTeclas(4, 9) - 349, ValorTopo(0) - 12)
            If ArrayTeclas(1, 9) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 9), 178, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 9)), Fonte, CorFonte, ArrayTeclas(5, 9) - 349, ValorTopo(1) - 12)
            If ArrayTeclas(0, 10) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 10), 186, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 10)), Fonte, CorFonte2, ArrayTeclas(4, 10) - 341, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 10) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 10), 186, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 10)), Fonte, CorFonte2, ArrayTeclas(5, 10) - 341, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 11) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 11), 195, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 11)), Fonte, CorFonte, ArrayTeclas(4, 11) - 332, ValorTopo(0) - 12)
            If ArrayTeclas(1, 11) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 11), 195, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 11)), Fonte, CorFonte, ArrayTeclas(5, 11) - 332, ValorTopo(1) - 12)
            If ArrayTeclas(0, 12) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 12), 203, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 12)), Fonte, CorFonte2, ArrayTeclas(4, 12) - 324, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 12) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 12), 203, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 12)), Fonte, CorFonte2, ArrayTeclas(5, 12) - 324, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 13) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 13), 212, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 13)), Fonte, CorFonte, ArrayTeclas(4, 13) - 315, ValorTopo(0) - 12)
            If ArrayTeclas(1, 13) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 13), 212, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 13)), Fonte, CorFonte, ArrayTeclas(5, 13) - 315, ValorTopo(1) - 12)
            If ArrayTeclas(0, 14) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 14), 220, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 14)), Fonte, CorFonte2, ArrayTeclas(4, 14) - 307, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 14) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 14), 220, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 14)), Fonte, CorFonte2, ArrayTeclas(5, 14) - 307, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 15) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 15), 229, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 15)), Fonte, CorFonte, ArrayTeclas(4, 15) - 298, ValorTopo(0) - 12)
            If ArrayTeclas(1, 15) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 15), 229, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 15)), Fonte, CorFonte, ArrayTeclas(5, 15) - 298, ValorTopo(1) - 12)
            If ArrayTeclas(0, 16) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 16), 246, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 16)), Fonte, CorFonte, ArrayTeclas(4, 16) - 281, ValorTopo(0) - 12)
            If ArrayTeclas(1, 16) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 16), 246, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 16)), Fonte, CorFonte, ArrayTeclas(5, 16) - 281, ValorTopo(1) - 12)
            If ArrayTeclas(0, 17) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 17), 254, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 17)), Fonte, CorFonte2, ArrayTeclas(4, 17) - 273, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 17) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 17), 254, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 17)), Fonte, CorFonte2, ArrayTeclas(5, 17) - 273, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 18) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 18), 263, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 18)), Fonte, CorFonte, ArrayTeclas(4, 18) - 264, ValorTopo(0) - 12)
            If ArrayTeclas(1, 18) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 18), 263, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 18)), Fonte, CorFonte, ArrayTeclas(5, 18) - 264, ValorTopo(1) - 12)
            If ArrayTeclas(0, 19) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 19), 271, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 19)), Fonte, CorFonte2, ArrayTeclas(4, 19) - 256, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 19) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 19), 271, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 19)), Fonte, CorFonte2, ArrayTeclas(5, 19) - 256, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 20) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 20), 280, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 20)), Fonte, CorFonte, ArrayTeclas(4, 20) - 247, ValorTopo(0) - 12)
            If ArrayTeclas(1, 20) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 20), 280, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 20)), Fonte, CorFonte, ArrayTeclas(5, 20) - 247, ValorTopo(1) - 12)
            If ArrayTeclas(0, 21) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 21), 297, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 21)), Fonte, CorFonte, ArrayTeclas(4, 21) - 230, ValorTopo(0) - 12)
            If ArrayTeclas(1, 21) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 21), 297, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 21)), Fonte, CorFonte, ArrayTeclas(5, 21) - 230, ValorTopo(1) - 12)
            If ArrayTeclas(0, 22) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 22), 305, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 22)), Fonte, CorFonte2, ArrayTeclas(4, 22) - 222, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 22) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 22), 305, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 22)), Fonte, CorFonte2, ArrayTeclas(5, 22) - 222, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 23) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 23), 314, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 23)), Fonte, CorFonte, ArrayTeclas(4, 23) - 213, ValorTopo(0) - 12)
            If ArrayTeclas(1, 23) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 23), 314, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 23)), Fonte, CorFonte, ArrayTeclas(5, 23) - 213, ValorTopo(1) - 12)
            If ArrayTeclas(0, 24) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 24), 322, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 24)), Fonte, CorFonte2, ArrayTeclas(4, 24) - 205, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 24) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 24), 322, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 24)), Fonte, CorFonte2, ArrayTeclas(5, 24) - 205, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 25) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 25), 331, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 25)), Fonte, CorFonte, ArrayTeclas(4, 25) - 196, ValorTopo(0) - 12)
            If ArrayTeclas(1, 25) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 25), 331, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 25)), Fonte, CorFonte, ArrayTeclas(5, 25) - 196, ValorTopo(1) - 12)
            If ArrayTeclas(0, 26) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 26), 339, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 26)), Fonte, CorFonte2, ArrayTeclas(4, 26) - 188, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 26) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 26), 339, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 26)), Fonte, CorFonte2, ArrayTeclas(5, 26) - 188, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 27) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 27), 348, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 27)), Fonte, CorFonte, ArrayTeclas(4, 27) - 179, ValorTopo(0) - 12)
            If ArrayTeclas(1, 27) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 27), 348, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 27)), Fonte, CorFonte, ArrayTeclas(5, 27) - 179, ValorTopo(1) - 12)
            If ArrayTeclas(0, 28) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 28), 365, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 28)), Fonte, CorFonte, ArrayTeclas(4, 28) - 162, ValorTopo(0) - 12)
            If ArrayTeclas(1, 28) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 28), 365, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 28)), Fonte, CorFonte, ArrayTeclas(5, 28) - 162, ValorTopo(1) - 12)
            If ArrayTeclas(0, 29) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 29), 373, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 29)), Fonte, CorFonte2, ArrayTeclas(4, 29) - 154, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 29) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 29), 373, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 29)), Fonte, CorFonte2, ArrayTeclas(5, 29) - 154, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 30) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 30), 382, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 30)), Fonte, CorFonte, ArrayTeclas(4, 30) - 145, ValorTopo(0) - 12)
            If ArrayTeclas(1, 30) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 30), 382, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 30)), Fonte, CorFonte, ArrayTeclas(5, 30) - 145, ValorTopo(1) - 12)
            If ArrayTeclas(0, 31) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 31), 390, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 31)), Fonte, CorFonte2, ArrayTeclas(4, 31) - 137, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 31) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 31), 390, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 31)), Fonte, CorFonte2, ArrayTeclas(5, 31) - 137, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 32) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 32), 399, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 32)), Fonte, CorFonte, ArrayTeclas(4, 32) - 128, ValorTopo(0) - 12)
            If ArrayTeclas(1, 32) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 32), 399, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 32)), Fonte, CorFonte, ArrayTeclas(5, 32) - 128, ValorTopo(1) - 12)
            If ArrayTeclas(0, 33) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 33), 416, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 33)), Fonte, CorFonte, ArrayTeclas(4, 33) - 111, ValorTopo(0) - 12)
            If ArrayTeclas(1, 33) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 33), 416, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 33)), Fonte, CorFonte, ArrayTeclas(5, 33) - 111, ValorTopo(1) - 12)
            If ArrayTeclas(0, 34) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 34), 424, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 34)), Fonte, CorFonte2, ArrayTeclas(4, 34) - 103, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 34) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 34), 424, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 34)), Fonte, CorFonte2, ArrayTeclas(5, 34) - 103, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 35) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 35), 433, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 35)), Fonte, CorFonte, ArrayTeclas(4, 35) - 94, ValorTopo(0) - 12)
            If ArrayTeclas(1, 35) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 35), 433, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 35)), Fonte, CorFonte, ArrayTeclas(5, 35) - 94, ValorTopo(1) - 12)
            If ArrayTeclas(0, 36) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 36), 441, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 36)), Fonte, CorFonte2, ArrayTeclas(4, 36) - 86, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 36) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 36), 441, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 36)), Fonte, CorFonte2, ArrayTeclas(5, 36) - 86, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 37) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 37), 450, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 37)), Fonte, CorFonte, ArrayTeclas(4, 37) - 77, ValorTopo(0) - 12)
            If ArrayTeclas(1, 37) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 37), 450, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 37)), Fonte, CorFonte, ArrayTeclas(5, 37) - 77, ValorTopo(1) - 12)
            If ArrayTeclas(0, 38) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 38), 458, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 38)), Fonte, CorFonte2, ArrayTeclas(4, 38) - 69, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 38) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 38), 458, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 38)), Fonte, CorFonte2, ArrayTeclas(5, 38) - 69, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 39) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 39), 467, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 39)), Fonte, CorFonte, ArrayTeclas(4, 39) - 60, ValorTopo(0) - 12)
            If ArrayTeclas(1, 39) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 39), 467, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 39)), Fonte, CorFonte, ArrayTeclas(5, 39) - 60, ValorTopo(1) - 12)
            If ArrayTeclas(0, 40) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 40), 484, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 40)), Fonte, CorFonte, ArrayTeclas(4, 40) - 43, ValorTopo(0) - 12)
            If ArrayTeclas(1, 40) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 40), 484, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 40)), Fonte, CorFonte, ArrayTeclas(5, 40) - 43, ValorTopo(1) - 12)
            If ArrayTeclas(0, 41) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 41), 492, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 41)), Fonte, CorFonte2, ArrayTeclas(4, 41) - 35, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 41) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 41), 492, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 41)), Fonte, CorFonte2, ArrayTeclas(5, 41) - 35, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 42) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 42), 501, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 42)), Fonte, CorFonte, ArrayTeclas(4, 42) - 26, ValorTopo(0) - 12)
            If ArrayTeclas(1, 42) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 42), 501, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 42)), Fonte, CorFonte, ArrayTeclas(5, 42) - 26, ValorTopo(1) - 12)
            If ArrayTeclas(0, 43) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 43), 509, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 43)), Fonte, CorFonte2, ArrayTeclas(4, 43) - 18, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 43) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 43), 509, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 43)), Fonte, CorFonte2, ArrayTeclas(5, 43) - 18, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 44) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 44), 518, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 44)), Fonte, CorFonte, ArrayTeclas(4, 44) - 9, ValorTopo(0) - 12)
            If ArrayTeclas(1, 44) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 44), 518, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 44)), Fonte, CorFonte, ArrayTeclas(5, 44) - 9, ValorTopo(1) - 12)
            If ArrayTeclas(0, 45) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 45), 535, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 45)), Fonte, CorFonte, ArrayTeclas(4, 45) + 8, ValorTopo(0) - 12)
            If ArrayTeclas(1, 45) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 45), 535, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 45)), Fonte, CorFonte, ArrayTeclas(5, 45) + 8, ValorTopo(1) - 12)
            If ArrayTeclas(0, 46) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 46), 543, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 46)), Fonte, CorFonte2, ArrayTeclas(4, 46) + 16, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 46) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 46), 543, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 46)), Fonte, CorFonte2, ArrayTeclas(5, 46) + 16, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 47) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 47), 552, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 47)), Fonte, CorFonte, ArrayTeclas(4, 47) + 25, ValorTopo(0) - 12)
            If ArrayTeclas(1, 47) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 47), 552, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 47)), Fonte, CorFonte, ArrayTeclas(5, 47) + 25, ValorTopo(1) - 12)
            If ArrayTeclas(0, 48) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 48), 560, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 48)), Fonte, CorFonte2, ArrayTeclas(4, 48) + 33, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 48) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 48), 560, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 48)), Fonte, CorFonte2, ArrayTeclas(5, 48) + 33, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 49) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 49), 569, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 49)), Fonte, CorFonte, ArrayTeclas(4, 49) + 42, ValorTopo(0) - 12)
            If ArrayTeclas(1, 49) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 49), 569, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 49)), Fonte, CorFonte, ArrayTeclas(5, 49) + 42, ValorTopo(1) - 12)
            If ArrayTeclas(0, 50) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 50), 577, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 50)), Fonte, CorFonte2, ArrayTeclas(4, 50) + 50, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 50) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 50), 577, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 50)), Fonte, CorFonte2, ArrayTeclas(5, 50) + 50, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 51) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 51), 586, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 51)), Fonte, CorFonte, ArrayTeclas(4, 51) + 59, ValorTopo(0) - 12)
            If ArrayTeclas(1, 51) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 51), 586, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 51)), Fonte, CorFonte, ArrayTeclas(5, 51) + 59, ValorTopo(1) - 12)
            If ArrayTeclas(0, 52) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 52), 603, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 52)), Fonte, CorFonte, ArrayTeclas(4, 52) + 76, ValorTopo(0) - 12)
            If ArrayTeclas(1, 52) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 52), 603, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 52)), Fonte, CorFonte, ArrayTeclas(5, 52) + 76, ValorTopo(1) - 12)
            If ArrayTeclas(0, 53) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 53), 611, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 53)), Fonte, CorFonte2, ArrayTeclas(4, 53) + 84, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 53) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 53), 611, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 53)), Fonte, CorFonte2, ArrayTeclas(5, 53) + 84, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 54) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 54), 620, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 54)), Fonte, CorFonte, ArrayTeclas(4, 54) + 93, ValorTopo(0) - 12)
            If ArrayTeclas(1, 54) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 54), 620, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 54)), Fonte, CorFonte, ArrayTeclas(5, 54) + 93, ValorTopo(1) - 12)
            If ArrayTeclas(0, 55) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 55), 628, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 55)), Fonte, CorFonte2, ArrayTeclas(4, 55) + 101, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 55) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 55), 628, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 55)), Fonte, CorFonte2, ArrayTeclas(5, 55) + 101, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 56) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 56), 637, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 56)), Fonte, CorFonte, ArrayTeclas(4, 56) + 110, ValorTopo(0) - 12)
            If ArrayTeclas(1, 56) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 56), 637, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 56)), Fonte, CorFonte, ArrayTeclas(5, 56) + 110, ValorTopo(1) - 12)
            If ArrayTeclas(0, 57) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 57), 654, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 57)), Fonte, CorFonte, ArrayTeclas(4, 57) + 127, ValorTopo(0) - 12)
            If ArrayTeclas(1, 57) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 57), 654, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 57)), Fonte, CorFonte, ArrayTeclas(5, 57) + 127, ValorTopo(1) - 12)
            If ArrayTeclas(0, 58) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 58), 662, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 58)), Fonte, CorFonte2, ArrayTeclas(4, 58) + 135, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 58) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 58), 662, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 58)), Fonte, CorFonte2, ArrayTeclas(5, 58) + 135, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 59) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 59), 671, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 59)), Fonte, CorFonte, ArrayTeclas(4, 59) + 144, ValorTopo(0) - 12)
            If ArrayTeclas(1, 59) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 59), 671, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 59)), Fonte, CorFonte, ArrayTeclas(5, 59) + 144, ValorTopo(1) - 12)
            If ArrayTeclas(0, 60) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 60), 679, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 60)), Fonte, CorFonte2, ArrayTeclas(4, 60) + 152, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 60) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 60), 679, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 60)), Fonte, CorFonte2, ArrayTeclas(5, 60) + 152, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 61) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 61), 688, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 61)), Fonte, CorFonte, ArrayTeclas(4, 61) + 161, ValorTopo(0) - 12)
            If ArrayTeclas(1, 61) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 61), 688, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 61)), Fonte, CorFonte, ArrayTeclas(5, 61) + 161, ValorTopo(1) - 12)
            If ArrayTeclas(0, 62) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 62), 696, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 62)), Fonte, CorFonte2, ArrayTeclas(4, 62) + 169, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 62) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 62), 696, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 62)), Fonte, CorFonte2, ArrayTeclas(5, 62) + 169, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 63) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 63), 705, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 63)), Fonte, CorFonte, ArrayTeclas(4, 63) + 178, ValorTopo(0) - 12)
            If ArrayTeclas(1, 63) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 63), 705, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 63)), Fonte, CorFonte, ArrayTeclas(5, 63) + 178, ValorTopo(1) - 12)
            If ArrayTeclas(0, 64) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 64), 722, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 64)), Fonte, CorFonte, ArrayTeclas(4, 64) + 195, ValorTopo(0) - 12)
            If ArrayTeclas(1, 64) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 64), 722, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 64)), Fonte, CorFonte, ArrayTeclas(5, 64) + 195, ValorTopo(1) - 12)
            If ArrayTeclas(0, 65) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 65), 730, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 65)), Fonte, CorFonte2, ArrayTeclas(4, 65) + 203, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 65) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 65), 730, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 65)), Fonte, CorFonte2, ArrayTeclas(5, 65) + 203, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 66) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 66), 739, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 66)), Fonte, CorFonte, ArrayTeclas(4, 66) + 212, ValorTopo(0) - 12)
            If ArrayTeclas(1, 66) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 66), 739, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 66)), Fonte, CorFonte, ArrayTeclas(5, 66) + 212, ValorTopo(1) - 12)
            If ArrayTeclas(0, 67) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 67), 747, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 67)), Fonte, CorFonte2, ArrayTeclas(4, 67) + 220, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 67) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 67), 747, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 67)), Fonte, CorFonte2, ArrayTeclas(5, 67) + 220, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 68) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 68), 756, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 68)), Fonte, CorFonte, ArrayTeclas(4, 68) + 229, ValorTopo(0) - 12)
            If ArrayTeclas(1, 68) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 68), 756, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 68)), Fonte, CorFonte, ArrayTeclas(5, 68) + 229, ValorTopo(1) - 12)
            If ArrayTeclas(0, 69) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 69), 773, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 69)), Fonte, CorFonte, ArrayTeclas(4, 69) + 246, ValorTopo(0) - 12)
            If ArrayTeclas(1, 69) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 69), 773, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 69)), Fonte, CorFonte, ArrayTeclas(5, 69) + 246, ValorTopo(1) - 12)
            If ArrayTeclas(0, 70) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 70), 781, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 70)), Fonte, CorFonte2, ArrayTeclas(4, 70) + 254, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 70) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 70), 781, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 70)), Fonte, CorFonte2, ArrayTeclas(5, 70) + 254, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 71) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 71), 790, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 71)), Fonte, CorFonte, ArrayTeclas(4, 71) + 263, ValorTopo(0) - 12)
            If ArrayTeclas(1, 71) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 71), 790, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 71)), Fonte, CorFonte, ArrayTeclas(5, 71) + 263, ValorTopo(1) - 12)
            If ArrayTeclas(0, 72) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 72), 798, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 72)), Fonte, CorFonte2, ArrayTeclas(4, 72) + 271, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 72) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 72), 798, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 72)), Fonte, CorFonte2, ArrayTeclas(5, 72) + 271, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 73) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 73), 807, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 73)), Fonte, CorFonte, ArrayTeclas(4, 73) + 280, ValorTopo(0) - 12)
            If ArrayTeclas(1, 73) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 73), 807, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 73)), Fonte, CorFonte, ArrayTeclas(5, 73) + 280, ValorTopo(1) - 12)
            If ArrayTeclas(0, 74) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 74), 815, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 74)), Fonte, CorFonte2, ArrayTeclas(4, 74) + 288, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 74) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 74), 815, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 74)), Fonte, CorFonte2, ArrayTeclas(5, 74) + 288, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 75) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 75), 824, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 75)), Fonte, CorFonte, ArrayTeclas(4, 75) + 297, ValorTopo(0) - 12)
            If ArrayTeclas(1, 75) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 75), 824, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 75)), Fonte, CorFonte, ArrayTeclas(5, 75) + 297, ValorTopo(1) - 12)
            If ArrayTeclas(0, 76) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 76), 841, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 76)), Fonte, CorFonte, ArrayTeclas(4, 76) + 314, ValorTopo(0) - 12)
            If ArrayTeclas(1, 76) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 76), 841, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 76)), Fonte, CorFonte, ArrayTeclas(5, 76) + 314, ValorTopo(1) - 12)
            If ArrayTeclas(0, 77) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 77), 849, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 77)), Fonte, CorFonte2, ArrayTeclas(4, 77) + 322, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 77) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 77), 849, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 77)), Fonte, CorFonte2, ArrayTeclas(5, 77) + 322, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 78) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 78), 858, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 78)), Fonte, CorFonte, ArrayTeclas(4, 78) + 331, ValorTopo(0) - 12)
            If ArrayTeclas(1, 78) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 78), 858, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 78)), Fonte, CorFonte, ArrayTeclas(5, 78) + 331, ValorTopo(1) - 12)
            If ArrayTeclas(0, 79) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 79), 866, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 79)), Fonte, CorFonte2, ArrayTeclas(4, 79) + 339, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 79) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 79), 866, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 79)), Fonte, CorFonte2, ArrayTeclas(5, 79) + 339, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 80) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 80), 875, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 80)), Fonte, CorFonte, ArrayTeclas(4, 80) + 348, ValorTopo(0) - 12)
            If ArrayTeclas(1, 80) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 80), 875, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 80)), Fonte, CorFonte, ArrayTeclas(5, 80) + 348, ValorTopo(1) - 12)
            If ArrayTeclas(0, 81) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 81), 892, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 81)), Fonte, CorFonte, ArrayTeclas(4, 81) + 365, ValorTopo(0) - 12)
            If ArrayTeclas(1, 81) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 81), 892, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 81)), Fonte, CorFonte, ArrayTeclas(5, 81) + 365, ValorTopo(1) - 12)
            If ArrayTeclas(0, 82) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 82), 900, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 82)), Fonte, CorFonte2, ArrayTeclas(4, 82) + 373, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 82) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 82), 900, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 82)), Fonte, CorFonte2, ArrayTeclas(5, 82) + 373, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 83) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 83), 909, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 83)), Fonte, CorFonte, ArrayTeclas(4, 83) + 382, ValorTopo(0) - 12)
            If ArrayTeclas(1, 83) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 83), 909, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 83)), Fonte, CorFonte, ArrayTeclas(5, 83) + 382, ValorTopo(1) - 12)
            If ArrayTeclas(0, 84) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 84), 917, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 84)), Fonte, CorFonte2, ArrayTeclas(4, 84) + 390, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 84) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 84), 917, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 84)), Fonte, CorFonte2, ArrayTeclas(5, 84) + 390, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 85) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 85), 926, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 85)), Fonte, CorFonte, ArrayTeclas(4, 85) + 399, ValorTopo(0) - 12)
            If ArrayTeclas(1, 85) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 85), 926, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 85)), Fonte, CorFonte, ArrayTeclas(5, 85) + 399, ValorTopo(1) - 12)
            If ArrayTeclas(0, 86) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 86), 934, ValorTopo(0) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 86)), Fonte, CorFonte2, ArrayTeclas(4, 86) + 407, ValorTopo(0) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(1, 86) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 86), 934, ValorTopo(1) - AjusteTeclasPretas, 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 86)), Fonte, CorFonte2, ArrayTeclas(5, 86) + 407, ValorTopo(1) - 12 - AjusteTeclasPretas)
            If ArrayTeclas(0, 87) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 87), 943, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 87)), Fonte, CorFonte, ArrayTeclas(4, 87) + 416, ValorTopo(0) - 12)
            If ArrayTeclas(1, 87) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 87), 943, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 87)), Fonte, CorFonte, ArrayTeclas(5, 87) + 416, ValorTopo(1) - 12)
            If ArrayTeclas(0, 88) > 0 Then gr.FillEllipse(CorKeyVelocity(0, 88), 960, ValorTopo(0), 12, 12) : gr.DrawString(CStr(ArrayTeclas(0, 88)), Fonte, CorFonte, ArrayTeclas(4, 88) + 433, ValorTopo(0) - 12)
            If ArrayTeclas(1, 88) > 0 Then gr.FillEllipse(CorKeyVelocity(1, 88), 960, ValorTopo(1), 12, 12) : gr.DrawString(CStr(ArrayTeclas(3, 88)), Fonte, CorFonte, ArrayTeclas(5, 88) + 433, ValorTopo(1) - 12)


            gr.FillRectangle(VPPP, 157, 208, 94, 46)
            gr.FillRectangle(VPP, 251, 208, 94, 46)
            gr.FillRectangle(VP, 345, 208, 94, 46)
            gr.FillRectangle(VMP, 438, 208, 94, 46)
            gr.FillRectangle(VMF, 532, 208, 94, 46)
            gr.FillRectangle(VF, 624, 208, 94, 46)
            gr.FillRectangle(VFF, 717, 208, 94, 46)
            gr.FillRectangle(VFFF, 811, 208, 94, 46)

            gr.SmoothingMode = SmoothingMode.None

            gr.DrawImage(My.Resources.LegendaDinâmicas, 97, 70, 868, 221)

            Me.SetBitmap(FaceBit, TransAmount)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub TreinamentoDinâmicaMãoEsquerdaDireita_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try

            ValorTopo(0) = 382
            ValorTopo(1) = 516
            AjusteTeclasPretas = 30

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

            Desenhar()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

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
                ArrayTeclas(0, [Param] - 20) = [KeyVelocity]
                ArrayTeclas(1, [Param] - 20) += [KeyVelocity]
                ArrayTeclas(2, [Param] - 20) += 1
                If ArrayTeclas(1, [Param] - 20) >= 0 Then ArrayTeclas(3, [Param] - 20) = CInt(ArrayTeclas(1, [Param] - 20) / ArrayTeclas(2, [Param] - 20)) 'calcula a média

                i = [Param] - 20

                Desenhar()

            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub DefineCor()

        Try

            If ArrayTeclas(0, i) >= 0 AndAlso ArrayTeclas(0, i) <= 15 Then
                CorKeyVelocity(0, i) = New SolidBrush(PPP.Color) 'amarelo claro
            ElseIf ArrayTeclas(0, i) >= 16 AndAlso ArrayTeclas(0, i) <= 31 Then
                CorKeyVelocity(0, i) = New SolidBrush(PP.Color) 'amarelo escuro
            ElseIf ArrayTeclas(0, i) >= 32 AndAlso ArrayTeclas(0, i) <= 47 Then
                CorKeyVelocity(0, i) = New SolidBrush(P.Color) 'laranja claro
            ElseIf ArrayTeclas(0, i) >= 48 AndAlso ArrayTeclas(0, i) <= 63 Then
                CorKeyVelocity(0, i) = New SolidBrush(MP.Color) 'laranja escuro
            ElseIf ArrayTeclas(0, i) >= 64 AndAlso ArrayTeclas(0, i) <= 79 Then
                CorKeyVelocity(0, i) = New SolidBrush(MF.Color) 'vermelho claro
            ElseIf ArrayTeclas(0, i) >= 80 AndAlso ArrayTeclas(0, i) <= 95 Then
                CorKeyVelocity(0, i) = New SolidBrush(F.Color) 'vermelho escuro
            ElseIf ArrayTeclas(0, i) >= 96 AndAlso ArrayTeclas(0, i) <= 111 Then
                CorKeyVelocity(0, i) = New SolidBrush(FF.Color) 'marrom
            ElseIf ArrayTeclas(0, i) >= 112 AndAlso ArrayTeclas(0, i) <= 127 Then
                CorKeyVelocity(0, i) = New SolidBrush(FFF.Color) 'marrom escuro
            End If
            If ArrayTeclas(3, i) >= 0 AndAlso ArrayTeclas(3, i) <= 15 Then
                CorKeyVelocity(1, i) = New SolidBrush(PPP.Color) 'amarelo claro
            ElseIf ArrayTeclas(3, i) >= 16 AndAlso ArrayTeclas(3, i) <= 31 Then
                CorKeyVelocity(1, i) = New SolidBrush(PP.Color) 'amarelo escuro
            ElseIf ArrayTeclas(3, i) >= 32 AndAlso ArrayTeclas(3, i) <= 47 Then
                CorKeyVelocity(1, i) = New SolidBrush(P.Color) 'laranja claro
            ElseIf ArrayTeclas(3, i) >= 48 AndAlso ArrayTeclas(3, i) <= 63 Then
                CorKeyVelocity(1, i) = New SolidBrush(MP.Color) 'laranja escuro
            ElseIf ArrayTeclas(3, i) >= 64 AndAlso ArrayTeclas(3, i) <= 79 Then
                CorKeyVelocity(1, i) = New SolidBrush(MF.Color) 'vermelho claro
            ElseIf ArrayTeclas(3, i) >= 80 AndAlso ArrayTeclas(3, i) <= 95 Then
                CorKeyVelocity(1, i) = New SolidBrush(F.Color) 'vermelho escuro
            ElseIf ArrayTeclas(3, i) >= 96 AndAlso ArrayTeclas(3, i) <= 111 Then
                CorKeyVelocity(1, i) = New SolidBrush(FF.Color) 'marrom
            ElseIf ArrayTeclas(3, i) >= 112 AndAlso ArrayTeclas(3, i) <= 127 Then
                CorKeyVelocity(1, i) = New SolidBrush(FFF.Color) 'marrom escuro
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub DefineCor2()

        Try

            For ii = 0 To 88
                If ArrayTeclas(0, ii) >= 0 AndAlso ArrayTeclas(0, ii) <= 15 Then
                    CorKeyVelocity(0, ii) = New SolidBrush(PPP.Color) 'amarelo claro
                ElseIf ArrayTeclas(0, ii) >= 16 AndAlso ArrayTeclas(0, ii) <= 31 Then
                    CorKeyVelocity(0, ii) = New SolidBrush(PP.Color) 'amarelo escuro
                ElseIf ArrayTeclas(0, ii) >= 32 AndAlso ArrayTeclas(0, ii) <= 47 Then
                    CorKeyVelocity(0, ii) = New SolidBrush(P.Color) 'laranja claro
                ElseIf ArrayTeclas(0, ii) >= 48 AndAlso ArrayTeclas(0, ii) <= 63 Then
                    CorKeyVelocity(0, ii) = New SolidBrush(MP.Color) 'laranja escuro
                ElseIf ArrayTeclas(0, ii) >= 64 AndAlso ArrayTeclas(0, ii) <= 79 Then
                    CorKeyVelocity(0, ii) = New SolidBrush(MF.Color) 'vermelho claro
                ElseIf ArrayTeclas(0, ii) >= 80 AndAlso ArrayTeclas(0, ii) <= 95 Then
                    CorKeyVelocity(0, ii) = New SolidBrush(F.Color) 'vermelho escuro
                ElseIf ArrayTeclas(0, ii) >= 96 AndAlso ArrayTeclas(0, ii) <= 111 Then
                    CorKeyVelocity(0, ii) = New SolidBrush(FF.Color) 'marrom
                ElseIf ArrayTeclas(0, ii) >= 112 AndAlso ArrayTeclas(0, ii) <= 127 Then
                    CorKeyVelocity(0, ii) = New SolidBrush(FFF.Color) 'marrom escuro
                End If
                If ArrayTeclas(3, ii) >= 0 AndAlso ArrayTeclas(3, ii) <= 15 Then
                    CorKeyVelocity(1, ii) = New SolidBrush(PPP.Color) 'amarelo claro
                ElseIf ArrayTeclas(3, ii) >= 16 AndAlso ArrayTeclas(3, ii) <= 31 Then
                    CorKeyVelocity(1, ii) = New SolidBrush(PP.Color) 'amarelo escuro
                ElseIf ArrayTeclas(3, ii) >= 32 AndAlso ArrayTeclas(3, ii) <= 47 Then
                    CorKeyVelocity(1, ii) = New SolidBrush(P.Color) 'laranja claro
                ElseIf ArrayTeclas(3, ii) >= 48 AndAlso ArrayTeclas(3, ii) <= 63 Then
                    CorKeyVelocity(1, ii) = New SolidBrush(MP.Color) 'laranja escuro
                ElseIf ArrayTeclas(3, ii) >= 64 AndAlso ArrayTeclas(3, ii) <= 79 Then
                    CorKeyVelocity(1, ii) = New SolidBrush(MF.Color) 'vermelho claro
                ElseIf ArrayTeclas(3, ii) >= 80 AndAlso ArrayTeclas(3, ii) <= 95 Then
                    CorKeyVelocity(1, ii) = New SolidBrush(F.Color) 'vermelho escuro
                ElseIf ArrayTeclas(3, ii) >= 96 AndAlso ArrayTeclas(3, ii) <= 111 Then
                    CorKeyVelocity(1, ii) = New SolidBrush(FF.Color) 'marrom
                ElseIf ArrayTeclas(3, ii) >= 112 AndAlso ArrayTeclas(3, ii) <= 127 Then
                    CorKeyVelocity(1, ii) = New SolidBrush(FFF.Color) 'marrom escuro
                End If
            Next
            SalvaSettings()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub TouchOff(ByVal [Param] As Byte, ByVal [KeyVelocity] As Byte)
        Try

            If Me.KeyCol([Param]).InvokeRequired Then
                Me.KeyCol([Param]).Invoke(DelgParamOff, New Object() {[Param], [KeyVelocity]})
            Else
                'Me.KeyCol([Param]).Visible = False

            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub ResetarArray(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click
        Try

            Array.Clear(ArrayTeclas, 0, ArrayTeclas.Length)
            Desenhar()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub ab_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown, Label1.MouseDown, CorPPP.MouseDown, CorPP.MouseDown, CorP.MouseDown, CorMP.MouseDown, CorMF.MouseDown, CorFFF.MouseDown, CorFF.MouseDown, CorF.MouseDown, CoresPadrão.MouseDown
        VGa = Me.MousePosition.X - Me.Location.X
        VGb = Me.MousePosition.Y - Me.Location.Y
    End Sub

    Private Sub ab_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove, Label1.MouseMove, CorPPP.MouseMove, CorPP.MouseMove, CorP.MouseMove, CorMP.MouseMove, CorMF.MouseMove, CorFFF.MouseMove, CorFF.MouseMove, CorF.MouseMove, CoresPadrão.MouseMove

        If e.Button = MouseButtons.Left Then
            VGnewPoint = Me.MousePosition
            VGnewPoint.X -= VGa
            VGnewPoint.Y -= VGb
            Me.Location = VGnewPoint
        End If

    End Sub

    Private Sub CorPPP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CorPPP.DoubleClick
        Try

            PPP.ShowDialog()
            My.Settings.NovoValorPPP = PPP.Color
            DefineCor2()
            Desenhar()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub CorPP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CorPP.DoubleClick
        Try

            PP.ShowDialog()
            My.Settings.NovoValorPP = PP.Color
            DefineCor2()
            Desenhar()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub CorP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CorP.DoubleClick
        Try

            P.ShowDialog()
            My.Settings.NovoValorP = P.Color
            DefineCor2()
            Desenhar()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub CorMP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CorMP.DoubleClick
        Try

            MP.ShowDialog()
            My.Settings.NovoValorMP = MP.Color
            DefineCor2()
            Desenhar()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub CorMF_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CorMF.DoubleClick
        Try

            MF.ShowDialog()
            My.Settings.NovoValorMF = MF.Color
            DefineCor2()
            Desenhar()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub CorF_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CorF.DoubleClick
        Try

            F.ShowDialog()
            My.Settings.NovoValorF = F.Color
            DefineCor2()
            Desenhar()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub CorFF_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CorFF.DoubleClick
        Try

            FF.ShowDialog()
            My.Settings.NovoValorFF = FF.Color
            DefineCor2()
            Desenhar()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub CorFFF_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CorFFF.DoubleClick
        Try

            FFF.ShowDialog()
            My.Settings.NovoValorFFF = FFF.Color
            DefineCor2()
            Desenhar()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Form_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try

            WinMM.midiInReset(hMidiIn) 'se não resetar o midiinclose não funcionará
            WinMM.midiInClose(hMidiIn)

            SalvaSettings()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Public Sub SalvaSettings()
        Try

            My.Settings.NovoValorPPP = PPP.Color
            My.Settings.NovoValorPP = PP.Color
            My.Settings.NovoValorP = P.Color
            My.Settings.NovoValorMP = MP.Color
            My.Settings.NovoValorMF = MF.Color
            My.Settings.NovoValorF = F.Color
            My.Settings.NovoValorFF = FF.Color
            My.Settings.NovoValorFFF = FFF.Color

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub CoresPadrão_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CoresPadrão.DoubleClick
        Try

            PPP.Color = Color.FromArgb(255, 255, 150)
            PP.Color = Color.FromArgb(255, 255, 0)
            P.Color = Color.FromArgb(255, 210, 0)
            MP.Color = Color.FromArgb(255, 150, 0)
            MF.Color = Color.FromArgb(255, 40, 40)
            F.Color = Color.FromArgb(215, 0, 0)
            FF.Color = Color.FromArgb(150, 50, 0)
            FFF.Color = Color.FromArgb(85, 0, 0)
            DefineCor2()
            Desenhar()

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
End Class