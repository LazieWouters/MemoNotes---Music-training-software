Option Strict On
Option Explicit On

Public Module VariáveisGlobais
    Public VGzz, NomeTimer, Tocanota, VGa, VGb, VGc, VGd, VGee, VGf, VGg, VGh, VGi, VGj, VGk, VGl, VGm, VGn, VGt, VGaa, qtdeloop, VGhhh, VGab, VGac, VGx, VGz, I_1_0, I_a1_1, I_b2_2, I_2_3, I_a2_4, I_b3_5, I_3_6, I_a3_7, I_b4_8, I_4_9, I_a4_10, I_b5_11, I_5_12, I_a5_13, I_b6_14, I_6_15, I_a6_16, I_b7_17, I_7_18, I_a7_19, I_b8_20, I_8_21, VGhMidiIn As Integer
    Public o, VGoo, VGrr, mp3, VGjjj, VGccc, NomeEscala, Intervalos, Intervalos2 As String
    Public ValorP(1, 87) As Integer
    Public a_I, a_II, a_III, a_IV, a_V, a_VI, a_VII As Integer
    Public NumeraçãoFamiliaAcorde As Integer
    Public VGnewPoint As New System.Drawing.Point()
    Public ToolTipNota2 As Bitmap
    Public pbox2 As PictureBox
    Private DecimaPrimeiraThread As Thread
    Public Percentual(,) As String = {{"", "", "", "", "", "", "x", "", "", "", "", "", "", "", "", "", "x", "", "", "", "", "", "", "", "x", "", "", "", "x", "", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "", "x", "", "x", "x", "x", "", "", "", "x", "", "", "", "x", "", "", "", "", "", "x", "", "x", "x", "x", "", "x", "x", "x", "", "x", "x", "", "x", "x", "x", "x", "x", "", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "x", "", "", "", "x", "", "x", "", "", "", "", "x", "x", "x", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "", "", "", "", "", "x", "x", "", "", "x", "", "", "", "", "", "x", "x", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "x", "", "", "", "", "", "", "", "", "", "", "", "x", "x", "x", "x", "x", "", "", "x", "x", "", "x", "x", "x", "x", "x", "x", "", "x", "", "x", "x", "", "", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "", "", "", "", "x", "", "", "", "", "x", "", "", "x", "", "x", "x", "", "", "x", "x", "x", "x", "", "x", "", "", "x", "x", "x", "x", "x", "x", "", "", "x", "x", "x", "", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "x", "", "", "x", "", "", "x", "x", "", "x", "", "", "x", "", "x", "", "", "x", "x", "", "", "", "", "x", "", "", "x", "", "x", "", "", "", "", "x", "x", "", "", "", "x", "x", "x", "x", "x", "x", "x", "x", "", "", "x", "x", "x", "x", "x", "x"}, _
{"", "", "x", "", "", "", "", "", "", "x", "", "x", "", "", "", "", "", "", "x", "x", "", "", "", "", "", "", "x", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "", "x", "", "", "x", "x", "x", "", "x", "", "", "x", "x", "x", "x", "x", "x", "", "x", "x"}, _
{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "x", "", "", "x", "x", "x", "", "", "", "", "x", "", "", "x", "x", "x", "", "", "", "x", "", "x", "x", "", "x", "x", "x", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "", "x", "x", "x", "x", "x"}, _
{"", "", "", "x", "", "x", "", "x", "", "", "", "", "x", "", "", "x", "", "", "x", "", "", "", "", "x", "x", "x", "", "", "x", "", "", "", "", "", "x", "", "", "x", "", "", "x", "x", "", "x", "x", "x", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "x", "", "x", "", "", "", "", "x", "", "", "x", "", "", "x", "x", "", "", "", "x", "", "x", "", "x", "x", "x", "x", "", "x", "", "x", "x", "x", "x", "x", "", "", "x", "x", "x", "", "", "x", "x", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "", "", "x", "", "", "", "", "", "", "", "x", "", "", "x", "x", "", "", "x", "x", "x", "", "", "", "", "x", "x", "x", "x", "", "x", "", "x", "x", "x", "x", "x", "", "x", "x", "", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "x", "", "", "", "", "", "", "", "", "", "", "x", "", "x", "", "", "", "", "x", "x", "x", "x", "", "", "", "x", "x", "x", "x", "", "", "x", "x", "x", "x", "", "", "", "x", "", "x", "x", "x", "x", "x", "x", "x", "", "x", "", "x", "x", "", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "", "", "", "", "", "x", "", "", "", "", "", "", "", "x", "", "", "", "", "x", "x", "x", "", "", "", "", "x", "", "", "x", "", "", "x", "x", "x", "", "", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "", "", "", "x", "", "", "", "x", "", "", "x", "x", "x", "", "", "x", "", "x", "", "", "", "", "x", "x", "x", "", "x", "", "", "x", "x", "", "x", "x", "x", "x", "x", "", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "x", "", "x", "x", "", "", "x", "", "", "x", "", "", "", "", "", "x", "", "", "x", "", "", "x", "", "x", "", "", "", "", "x", "x", "", "x", "x", "", "", "x", "", "x", "x", "x", "", "x", "x", "x", "x", "", "x", "", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "x", "", "", "", "x", "", "", "x", "", "x", "x", "x", "x", "x", "", "x", "", "", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "", "x", "", "x", "", "x", "", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "x", "", "", "", "", "", "", "", "x", "x", "", "x", "x", "", "x", "", "", "", "", "", "", "", "", "x", "", "x", "x", "x", "", "x", "", "", "", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "x", "", "", "", "", "", "", "", "x", "", "x", "x", "x", "", "", "x", "", "x", "", "x", "", "x", "", "x", "x", "x", "x", "x", "x", "", "", "x", "x", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "", "", "", "", "x", "", "", "", "", "", "", "x", "", "x", "", "", "x", "", "", "", "x", "x", "", "x", "x", "x", "", "x", "", "x", "x", "x", "", "x", "x", "", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "", "", "", "x", "", "", "", "", "", "", "x", "", "", "", "", "", "x", "x", "", "x", "", "", "", "x", "", "x", "x", "x", "x", "x", "x", "", "x", "", "", "x", "", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "", "x"}, _
{"", "", "", "", "", "", "x", "x", "", "x", "x", "", "", "", "x", "x", "", "x", "", "", "x", "x", "x", "", "", "", "x", "", "", "x", "x", "", "", "", "", "", "x", "", "x", "x", "x", "x", "x", "", "x", "", "x", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "", "", "", "", "x", "", "", "", "x", "", "", "", "", "", "", "x", "", "", "", "", "x", "", "x", "x", "x", "x", "x", "x", "", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "x", "x", "", "", "", "x", "x", "x", "x", "", "", "", "x", "", "", "x", "", "", "", "x", "", "", "x", "x", "x", "x", "x", "", "", "x", "x", "x", "x", "x", "x", "", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "", "", "", "", "", "x", "", "", "x", "x", "x", "", "", "x", "", "x", "", "", "", "", "", "x", "x", "", "x", "x", "", "", "x", "x", "x", "x", "x", "", "x", "", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "", "", "", "x", "", "", "x", "", "", "", "", "", "", "", "x", "", "", "x", "x", "", "x", "", "", "x", "", "", "x", "x", "", "", "", "", "x", "x", "", "x", "x", "x", "x", "x", "", "x", "x", "", "", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "x", "x", "x", "", "x", "x", "", "", "x", "", "x", "x", "x", "", "", "", "x", "x", "", "", "x", "x", "x", "x", "x", "", "x", "x", "x", "x", "", "x", "x", "x", "x", "x", "", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "", "", "", "", "x", "x", "", "", "", "x", "", "", "", "", "", "", "", "x", "x", "x", "x", "x", "", "x", "x", "x", "x", "x", "", "x", "x", "x", "", "", "", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "x", "", "", "", "", "", "x", "", "", "", "", "x", "", "", "", "", "x", "", "", "x", "", "", "x", "x", "x", "", "", "", "x", "", "", "", "x", "x", "", "", "x", "x", "x", "x", "x", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "x", "", "", "", "", "", "", "x", "", "", "", "x", "", "x", "", "", "x", "x", "x", "", "x", "x", "", "x", "x", "", "", "", "x", "x", "x", "x", "x", "x", "x", "", "", "x", "x", "x", "", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "", "", "", "x", "", "", "", "", "", "", "", "", "", "", "", "x", "", "", "", "x", "x", "", "", "", "", "", "", "x", "x", "x", "x", "x", "", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "", "x", "", "", "", "", "", "x", "", "", "", "", "x", "x", "x", "", "", "x", "x", "", "", "x", "", "x", "", "x", "x", "", "", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "", "x", "x", "x", "x"}, _
{"", "", "", "x", "", "", "", "", "", "", "", "", "", "x", "x", "", "x", "x", "", "", "", "", "", "", "x", "", "", "x", "", "", "x", "x", "x", "", "x", "x", "x", "", "x", "", "", "x", "", "x", "", "x", "x", "", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "x", "x", "", "", "", "", "x", "", "", "", "", "", "x", "x", "", "x", "", "x", "x", "x", "x", "", "x", "", "x", "", "x", "x", "x", "", "", "x", "x", "", "x", "x", "", "x", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "x", "", "", "", "x", "", "x", "x", "", "", "", "x", "", "", "x", "", "", "", "x", "x", "x", "x", "x", "x", "", "", "x", "", "", "", "x", "x", "x", "x", "x", "x", "", "x", "x", "x", "", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "", "x", "", "", "", "x", "", "", "x", "", "", "", "x", "x", "", "", "x", "x", "", "", "", "", "", "", "x", "x", "x", "x", "x", "x", "x", "", "x", "x", "x", "", "x", "", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "x", "", "", "", "x", "x", "", "x", "x", "", "", "", "", "x", "", "", "x", "x", "x", "x", "x", "x", "x", "", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "", "", "", "x", "", "", "", "", "", "", "x", "x", "x", "", "x", "x", "", "", "x", "", "x", "x", "", "", "x", "x", "x", "x", "", "x", "x", "x", "x", "", "", "x", "", "x", "x", "", "x", "x", "", "x", "x", "x", "x", "", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "x", "", "", "", "", "", "", "", "", "x", "", "x", "", "", "", "", "", "", "", "x", "", "x", "x", "", "x", "", "", "x", "", "x", "x", "", "", "", "x", "x", "x", "x", "x", "x", "", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "x", "", "", "", "x", "", "", "", "", "x", "", "x", "", "", "", "", "", "", "", "x", "", "x", "", "", "x", "", "x", "", "", "x", "", "", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "x", "", "", "", "", "", "", "", "", "", "", "", "", "", "x", "", "x", "", "", "", "", "", "x", "x", "x", "x", "x", "", "x", "x", "x", "x", "", "x", "x", "x", "", "", "", "x", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "x", "", "", "", "", "", "", "", "", "x", "", "", "x", "", "", "", "", "x", "", "", "", "x", "x", "", "x", "x", "x", "x", "x", "", "x", "", "x", "", "x", "x", "x", "x", "x", "", "", "x", "x", "x", "", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "", "", "x", "", "x", "x", "x", "", "", "", "", "", "", "x", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "x", "", "", "", "", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "x", "x", "", "x", "x", "", "", "x", "", "", "x", "x", "", "x", "", "", "", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "", "x", "x", "", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "", "", "", "", "", "x", "", "", "", "", "", "", "", "x", "x", "", "x", "", "", "", "", "x", "", "x", "x", "", "x", "x", "", "x", "x", "x", "x", "x", "", "x", "", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "", "x", "", "x", "x"}, _
{"", "", "", "", "", "", "", "", "x", "", "", "", "", "", "", "", "", "", "", "", "x", "", "", "x", "", "", "", "x", "", "", "x", "", "x", "x", "", "", "x", "", "", "", "", "", "x", "x", "x", "x", "x", "x", "x", "x", "", "", "x", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "x", "", "", "", "x", "", "", "", "", "", "", "x", "", "", "x", "", "x", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "", "", "x", "", "", "", "", "", "x", "x", "", "", "", "", "x", "", "", "x", "", "", "x", "", "", "", "", "x", "", "x", "x", "", "x", "x", "x", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "", "x", "", "", "x", "", "", "", "", "", "x", "", "", "x", "", "", "x", "", "x", "x", "", "x", "x", "x", "x", "", "x", "x", "x", "x", "x", "", "", "x", "", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "x", "", "", "", "", "", "", "x", "", "x", "", "", "", "x", "", "", "x", "", "", "", "", "x", "", "", "", "x", "x", "", "x", "", "x", "", "x", "x", "x", "x", "x", "x", "", "", "x", "", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "x", "", "", "", "", "", "", "x", "", "", "", "", "", "", "", "", "", "", "x", "", "x", "x", "x", "x", "", "x", "x", "", "", "x", "x", "", "x", "", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "x", "", "", "", "", "", "", "x", "", "", "", "x", "x", "x", "", "", "", "", "", "", "", "", "x", "", "x", "", "", "", "", "x", "x", "x", "x", "", "", "x", "", "", "x", "", "x", "", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "x", "", "", "", "", "", "", "x", "", "", "x", "", "", "", "", "x", "", "x", "x", "x", "x", "", "x", "x", "x", "", "x", "x", "x", "", "x", "", "x", "", "x", "", "x", "x", "x", "", "x", "", "x", "x", "x", "x", "", "x", "", "x", "x", "x", "", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "x", "", "", "", "x", "", "", "", "", "x", "", "", "", "x", "x", "x", "", "x", "", "", "", "x", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "x", "", "", "", "", "x", "", "", "", "", "", "", "", "", "", "x", "x", "", "", "x", "", "x", "", "", "", "", "x", "", "", "", "x", "x", "", "x", "x", "", "x", "x", "", "x", "", "x", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "", "", "x", "", "", "", "x", "x", "x", "", "", "", "x", "x", "", "x", "", "", "x", "x", "x", "x", "x", "", "", "x", "", "x", "x", "", "", "", "x", "x", "x", "x", "x", "", "", "x", "x", "", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "", "x", "", "", "x", "", "", "", "", "", "", "", "", "", "", "", "", "x", "", "", "", "", "x", "x", "x", "", "x", "x", "", "x", "x", "x", "", "", "", "x", "", "x", "x", "", "x", "x", "", "x", "x", "x", "x", "x", "", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "x", "x", "x", "", "", "x", "x", "x", "", "x", "", "x", "", "x", "", "x", "", "", "", "x", "", "", "x", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "x", "", "x", "", "", "", "", "", "", "x", "x", "x", "", "x", "", "x", "x", "", "", "", "x", "x", "x", "x", "", "", "", "", "x", "x", "x", "", "x", "x", "x", "x", "", "x", "x", "x", "x", "x", "", "x", "", "x", "", "x", "x", "x", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "x", "", "", "x", "", "x", "x", "", "", "", "", "", "", "", "x", "", "", "", "", "", "", "x", "x", "x", "x", "", "", "x", "", "", "", "", "", "x", "", "x", "x", "", "", "x", "x", "x", "x", "x", "x", "", "", "x", "x", "x", "x", "x", "x"}, _
{"", "", "", "", "", "", "", "", "", "", "x", "", "", "", "", "", "", "", "", "x", "x", "x", "", "x", "x", "x", "", "", "x", "x", "", "", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "", "x", "x", "x", "", "x", "x", "x", "x", "x", "x", "x", "", "x", "x", "x", "x", "x"}}

    Public Sub ToolTipNotas()
        ToolTipNota2 = Nothing

        If pbox2.Name = "P1" OrElse pbox2.Name = "P13" OrElse pbox2.Name = "P25" OrElse pbox2.Name = "P37" OrElse pbox2.Name = "P49" OrElse pbox2.Name = "P61" OrElse pbox2.Name = "P73" OrElse pbox2.Name = "P85" Then
            ToolTipNota2 = My.Resources.ToolTip_A
        ElseIf pbox2.Name = "P2" OrElse pbox2.Name = "P14" OrElse pbox2.Name = "P26" OrElse pbox2.Name = "P38" OrElse pbox2.Name = "P50" OrElse pbox2.Name = "P62" OrElse pbox2.Name = "P74" OrElse pbox2.Name = "P86" Then
            ToolTipNota2 = My.Resources.ToolTip_Asus_Bb
        ElseIf pbox2.Name = "P3" OrElse pbox2.Name = "P15" OrElse pbox2.Name = "P27" OrElse pbox2.Name = "P39" OrElse pbox2.Name = "P51" OrElse pbox2.Name = "P63" OrElse pbox2.Name = "P75" OrElse pbox2.Name = "P87" Then
            ToolTipNota2 = My.Resources.ToolTip_B_Cb
        ElseIf pbox2.Name = "P4" OrElse pbox2.Name = "P16" OrElse pbox2.Name = "P28" OrElse pbox2.Name = "P40" OrElse pbox2.Name = "P52" OrElse pbox2.Name = "P64" OrElse pbox2.Name = "P76" OrElse pbox2.Name = "P88" Then
            ToolTipNota2 = My.Resources.ToolTip_C_Bsus
        ElseIf pbox2.Name = "P5" OrElse pbox2.Name = "P17" OrElse pbox2.Name = "P29" OrElse pbox2.Name = "P41" OrElse pbox2.Name = "P53" OrElse pbox2.Name = "P65" OrElse pbox2.Name = "P77" Then
            ToolTipNota2 = My.Resources.ToolTip_Csus_Db
        ElseIf pbox2.Name = "P6" OrElse pbox2.Name = "P18" OrElse pbox2.Name = "P30" OrElse pbox2.Name = "P42" OrElse pbox2.Name = "P54" OrElse pbox2.Name = "P66" OrElse pbox2.Name = "P78" Then
            ToolTipNota2 = My.Resources.ToolTip_D
        ElseIf pbox2.Name = "P7" OrElse pbox2.Name = "P19" OrElse pbox2.Name = "P31" OrElse pbox2.Name = "P43" OrElse pbox2.Name = "P55" OrElse pbox2.Name = "P67" OrElse pbox2.Name = "P79" Then
            ToolTipNota2 = My.Resources.ToolTip_Dsus_Eb
        ElseIf pbox2.Name = "P8" OrElse pbox2.Name = "P20" OrElse pbox2.Name = "P32" OrElse pbox2.Name = "P44" OrElse pbox2.Name = "P56" OrElse pbox2.Name = "P68" OrElse pbox2.Name = "P80" Then
            ToolTipNota2 = My.Resources.ToolTip_E_Fb
        ElseIf pbox2.Name = "P9" OrElse pbox2.Name = "P21" OrElse pbox2.Name = "P33" OrElse pbox2.Name = "P45" OrElse pbox2.Name = "P57" OrElse pbox2.Name = "P69" OrElse pbox2.Name = "P81" Then
            ToolTipNota2 = My.Resources.ToolTip_F_Esus
        ElseIf pbox2.Name = "P10" OrElse pbox2.Name = "P22" OrElse pbox2.Name = "P34" OrElse pbox2.Name = "P46" OrElse pbox2.Name = "P58" OrElse pbox2.Name = "P70" OrElse pbox2.Name = "P82" Then
            ToolTipNota2 = My.Resources.ToolTip_Fsus_Gb
        ElseIf pbox2.Name = "P11" OrElse pbox2.Name = "P23" OrElse pbox2.Name = "P35" OrElse pbox2.Name = "P47" OrElse pbox2.Name = "P59" OrElse pbox2.Name = "P71" OrElse pbox2.Name = "P83" Then
            ToolTipNota2 = My.Resources.ToolTip_G
        ElseIf pbox2.Name = "P12" OrElse pbox2.Name = "P24" OrElse pbox2.Name = "P36" OrElse pbox2.Name = "P48" OrElse pbox2.Name = "P60" OrElse pbox2.Name = "P72" OrElse pbox2.Name = "P84" Then
            ToolTipNota2 = My.Resources.ToolTip_Gsus_Ab
        ElseIf pbox2.Name = "CorTecla" OrElse pbox2.Name = "CorTecla2" Then
            tooltipNotas2()
        End If
    End Sub

    Public Sub tooltipNotas2()
        If o = "A0" OrElse o = "A1" OrElse o = "A2" OrElse o = "A3" OrElse o = "A4" OrElse o = "A5" OrElse o = "A6" OrElse o = "A7" Then
            ToolTipNota2 = My.Resources.ToolTip_A
        ElseIf o = "B0" OrElse o = "B1" OrElse o = "B2" OrElse o = "B3" OrElse o = "B4" OrElse o = "B5" OrElse o = "B6" OrElse o = "B7" OrElse o = "Cb1" OrElse o = "Cb2" OrElse o = "Cb3" OrElse o = "Cb4" OrElse o = "Cb5" OrElse o = "Cb6" OrElse o = "Cb7" OrElse o = "Cb8" Then
            ToolTipNota2 = My.Resources.ToolTip_B_Cb
        ElseIf o = "C1" OrElse o = "C2" OrElse o = "C3" OrElse o = "C4" OrElse o = "C5" OrElse o = "C6" OrElse o = "C7" OrElse o = "C8" OrElse o = "B#0" OrElse o = "B#1" OrElse o = "B#2" OrElse o = "B#3" OrElse o = "B#4" OrElse o = "B#5" OrElse o = "B#6" OrElse o = "B#7" Then
            ToolTipNota2 = My.Resources.ToolTip_C_Bsus
        ElseIf o = "D1" OrElse o = "D2" OrElse o = "D3" OrElse o = "D4" OrElse o = "D5" OrElse o = "D6" OrElse o = "D7" Then
            ToolTipNota2 = My.Resources.ToolTip_D
        ElseIf o = "E1" OrElse o = "E2" OrElse o = "E3" OrElse o = "E4" OrElse o = "E5" OrElse o = "E6" OrElse o = "E7" OrElse o = "Fb1" OrElse o = "Fb2" OrElse o = "Fb3" OrElse o = "Fb4" OrElse o = "Fb5" OrElse o = "Fb6" OrElse o = "Fb7" Then
            ToolTipNota2 = My.Resources.ToolTip_E_Fb
        ElseIf o = "F1" OrElse o = "F2" OrElse o = "F3" OrElse o = "F4" OrElse o = "F5" OrElse o = "F6" OrElse o = "F7" OrElse o = "E#1" OrElse o = "E#2" OrElse o = "E#3" OrElse o = "E#4" OrElse o = "E#5" OrElse o = "E#6" OrElse o = "E#7" Then
            ToolTipNota2 = My.Resources.ToolTip_F_Esus
        ElseIf o = "G1" OrElse o = "G2" OrElse o = "G3" OrElse o = "G4" OrElse o = "G5" OrElse o = "G6" OrElse o = "G7" Then
            ToolTipNota2 = My.Resources.ToolTip_G
        ElseIf o = "A#0" OrElse o = "A#1" OrElse o = "A#2" OrElse o = "A#3" OrElse o = "A#4" OrElse o = "A#5" OrElse o = "A#6" OrElse o = "A#7" OrElse o = "Bb0" OrElse o = "Bb1" OrElse o = "Bb2" OrElse o = "Bb3" OrElse o = "Bb4" OrElse o = "Bb5" OrElse o = "Bb6" OrElse o = "Bb7" Then
            ToolTipNota2 = My.Resources.ToolTip_Asus_Bb
        ElseIf o = "C#1" OrElse o = "C#2" OrElse o = "C#3" OrElse o = "C#4" OrElse o = "C#5" OrElse o = "C#6" OrElse o = "C#7" OrElse o = "C#8" OrElse o = "Db1" OrElse o = "Db2" OrElse o = "Db3" OrElse o = "Db4" OrElse o = "Db5" OrElse o = "Db6" OrElse o = "Db7" Then
            ToolTipNota2 = My.Resources.ToolTip_Csus_Db
        ElseIf o = "D#1" OrElse o = "D#2" OrElse o = "D#3" OrElse o = "D#4" OrElse o = "D#5" OrElse o = "D#6" OrElse o = "D#7" OrElse o = "Eb1" OrElse o = "Eb2" OrElse o = "Eb3" OrElse o = "Eb4" OrElse o = "Eb5" OrElse o = "Eb6" OrElse o = "Eb7" Then
            ToolTipNota2 = My.Resources.ToolTip_Dsus_Eb
        ElseIf o = "F#1" OrElse o = "F#2" OrElse o = "F#3" OrElse o = "F#4" OrElse o = "F#5" OrElse o = "F#6" OrElse o = "F#7" OrElse o = "Gb1" OrElse o = "Gb2" OrElse o = "Gb3" OrElse o = "Gb4" OrElse o = "Gb5" OrElse o = "Gb6" OrElse o = "Gb7" Then
            ToolTipNota2 = My.Resources.ToolTip_Fsus_Gb
        ElseIf o = "G#1" OrElse o = "G#2" OrElse o = "G#3" OrElse o = "G#4" OrElse o = "G#5" OrElse o = "G#6" OrElse o = "G#7" OrElse o = "Ab0" OrElse o = "Ab1" OrElse o = "Ab2" OrElse o = "Ab3" OrElse o = "Ab4" OrElse o = "Ab5" OrElse o = "Ab6" OrElse o = "Ab7" Then
            ToolTipNota2 = My.Resources.ToolTip_Gsus_Ab
        End If
    End Sub

End Module