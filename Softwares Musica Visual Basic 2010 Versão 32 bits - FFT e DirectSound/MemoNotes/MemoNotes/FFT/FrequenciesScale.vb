Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms
Imports System.Drawing

Partial Public Class FrequenciesScale
    Inherits UserControl
    Dim MinFrequency As Double = 40 '65 '70
    Dim MaxFrequency As Double = 4000 '1200
    Const AFrequency As Double = 440
    Shared ToneStep As Double = Math.Pow(2, 1.0 / 12)
    Dim Fonte As Font = New Font("Arial", 6)

    Shared Labels As ScaleLabel() = {New ScaleLabel() With { _
      .Title = "E1", _
     .Frequency = 41.2034, _
     .Color = Color.Transparent _
    }, New ScaleLabel() With { _
      .Title = "F1", _
     .Frequency = 43.6535, _
     .Color = Color.Transparent _
    }, New ScaleLabel() With { _
      .Title = "G1", _
     .Frequency = 48.9994, _
     .Color = Color.Transparent _
    }, New ScaleLabel() With { _
      .Title = "A1", _
     .Frequency = 55.0, _
     .Color = Color.Transparent _
    }, New ScaleLabel() With { _
      .Title = "B1", _
     .Frequency = 61.7354, _
     .Color = Color.Transparent _
    }, New ScaleLabel() With { _
     .Title = "C2", _
     .Frequency = 65.4064, _
     .Color = Color.Transparent _
    }, New ScaleLabel() With { _
     .Title = "D2", _
     .Frequency = 73.4162, _
     .Color = Color.Transparent _
    }, New ScaleLabel() With { _
     .Title = "E2", _
     .Frequency = 82.4069, _
     .Color = Color.DeepSkyBlue _
  }, New ScaleLabel() With { _
     .Title = "F2", _
     .Frequency = 87.3071, _
     .Color = Color.Transparent _
  }, New ScaleLabel() With { _
     .Title = "G2", _
     .Frequency = 97.9989, _
     .Color = Color.Transparent _
   }, New ScaleLabel() With { _
     .Title = "A2", _
     .Frequency = 110.0, _
     .Color = Color.DeepSkyBlue _
  }, New ScaleLabel() With { _
     .Title = "B2", _
     .Frequency = 123.4708, _
     .Color = Color.Transparent _
 }, New ScaleLabel() With { _
     .Title = "C3", _
     .Frequency = 130.8128, _
     .Color = Color.Transparent _
    }, New ScaleLabel() With { _
     .Title = "D3", _
     .Frequency = 146.8324, _
     .Color = Color.DeepSkyBlue _
 }, New ScaleLabel() With { _
     .Title = "E3", _
     .Frequency = 164.8138, _
     .Color = Color.Transparent _
 }, New ScaleLabel() With { _
     .Title = "F3", _
     .Frequency = 174.6141, _
     .Color = Color.Transparent _
    }, New ScaleLabel() With { _
     .Title = "G3", _
     .Frequency = 195.9977, _
     .Color = Color.DeepSkyBlue _
 }, New ScaleLabel() With { _
     .Title = "A3", _
     .Frequency = 220.0, _
     .Color = Color.Transparent _
    }, New ScaleLabel() With { _
     .Title = "B3", _
     .Frequency = 246.9417, _
     .Color = Color.DeepSkyBlue _
     }, New ScaleLabel() With { _
     .Title = "C4", _
     .Frequency = 261.6256, _
     .Color = Color.Red _
  }, New ScaleLabel() With { _
     .Title = "D4", _
     .Frequency = 293.6648, _
     .Color = Color.Transparent _
    }, New ScaleLabel() With { _
     .Title = "E4", _
     .Frequency = 329.6276, _
     .Color = Color.DeepSkyBlue _
     }, New ScaleLabel() With { _
     .Title = "F4", _
     .Frequency = 349.2282, _
     .Color = Color.Transparent _
      }, New ScaleLabel() With { _
     .Title = "G4", _
     .Frequency = 391.9954, _
     .Color = Color.Transparent _
    }, New ScaleLabel() With { _
     .Title = "A4", _
     .Frequency = 440.0, _
     .Color = Color.Red _
  }, New ScaleLabel() With { _
     .Title = "B4", _
     .Frequency = 493.8833, _
     .Color = Color.Transparent _
     }, New ScaleLabel() With { _
     .Title = "C5", _
     .Frequency = 523.2511, _
     .Color = Color.Transparent _
    }, New ScaleLabel() With { _
     .Title = "D5", _
     .Frequency = 587.3295, _
     .Color = Color.Transparent _
 }, New ScaleLabel() With { _
     .Title = "E5", _
     .Frequency = 659.2551, _
     .Color = Color.Transparent _
 }, New ScaleLabel() With { _
     .Title = "F5", _
     .Frequency = 698.4565, _
     .Color = Color.Transparent _
    }, New ScaleLabel() With { _
     .Title = "G5", _
     .Frequency = 783.9909, _
     .Color = Color.Transparent _
 }, New ScaleLabel() With { _
     .Title = "A5", _
     .Frequency = 880.0, _
     .Color = Color.Transparent _
    }, New ScaleLabel() With { _
     .Title = "B5", _
     .Frequency = 987.7666, _
     .Color = Color.Transparent _
  }, New ScaleLabel() With { _
     .Title = "C6", _
     .Frequency = 1046.5, _
     .Color = Color.Transparent _
    }, New ScaleLabel() With { _
     .Title = "D6", _
     .Frequency = 1174.66, _
     .Color = Color.Transparent _
    }, New ScaleLabel() With { _
     .Title = "E6", _
     .Frequency = 1318.51, _
     .Color = Color.Transparent _
     }, New ScaleLabel() With { _
     .Title = "F6", _
     .Frequency = 1396.91, _
     .Color = Color.Transparent _
     }, New ScaleLabel() With { _
     .Title = "G6", _
     .Frequency = 1567.98, _
     .Color = Color.Transparent _
     }, New ScaleLabel() With { _
     .Title = "A6", _
     .Frequency = 1760.0, _
     .Color = Color.Transparent _
     }, New ScaleLabel() With { _
     .Title = "B6", _
     .Frequency = 1975.53, _
     .Color = Color.Transparent _
     }, New ScaleLabel() With { _
     .Title = "C7", _
     .Frequency = 2093.0, _
     .Color = Color.Transparent _
     }, New ScaleLabel() With { _
     .Title = "D7", _
     .Frequency = 2349.32, _
     .Color = Color.Transparent _
      }, New ScaleLabel() With { _
     .Title = "E7", _
     .Frequency = 2637.02, _
     .Color = Color.Transparent _
      }, New ScaleLabel() With { _
     .Title = "F7", _
     .Frequency = 2793.83, _
     .Color = Color.Transparent _
      }, New ScaleLabel() With { _
     .Title = "G7", _
     .Frequency = 3135.96, _
     .Color = Color.Transparent _
      }, New ScaleLabel() With { _
     .Title = "A7", _
     .Frequency = 3520.0, _
     .Color = Color.Transparent _
      }, New ScaleLabel() With { _
     .Title = "B7", _
     .Frequency = 3951.07, _
     .Color = Color.Transparent _
     }, New ScaleLabel() With { _
     .Title = "C8", _
     .Frequency = 4186.01, _
     .Color = Color.Transparent _
    }}


    Private m_currentFrequency As Double

    <DefaultValue(0.0)> _
    Public Property CurrentFrequency() As Double
        Get
            Return m_currentFrequency
        End Get
        Set(ByVal value As Double)
            If m_currentFrequency <> value Then
                m_currentFrequency = value
                Invalidate()
            End If
        End Set
    End Property

    Private m_closestFrequency As Double

    <DefaultValue(0.0)> _
    Public Property ClosestFrequency() As Double
        Get
            Return m_closestFrequency
        End Get
        Set(ByVal value As Double)
            If m_closestFrequency <> value Then
                m_closestFrequency = value
                Invalidate()
            End If
        End Set
    End Property

    Private m_noteName As String

    <DefaultValue(0.0)> _
    Public Property NoteName() As String
        Get
            Return m_noteName
        End Get
        Set(ByVal value As String)
            If m_noteName <> value Then
                m_noteName = value
                Invalidate()
            End If
        End Set
    End Property

    Private m_signalDetected As Boolean = False

    <DefaultValue(False)> _
    Public Property SignalDetected() As Boolean
        Get
            Return m_signalDetected
        End Get
        Set(ByVal value As Boolean)
            If m_signalDetected <> value Then
                m_signalDetected = value
                Invalidate()
            End If
        End Set
    End Property

    Public Sub New()
        InitializeComponent()

        InitializeComponent2()
    End Sub

    Public Sub New(ByVal container As IContainer)
        container.Add(Me)

        InitializeComponent()

        InitializeComponent2()
    End Sub

    Private Sub InitializeComponent2()
        SetStyle(ControlStyles.ResizeRedraw, True)
        SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
    End Sub

    Shared MarkerPen As New Pen(Color.White)
    Shared MarkerPen2 As New Pen(Color.White, 3)
    Shared ActiveSliderBrush1 As Brush = New SolidBrush(Color.GreenYellow)
    Shared ActiveSliderBrush2 As Brush = New SolidBrush(Color.Green)
    Shared ActiveSliderBrush1B As Brush = New SolidBrush(Color.FromArgb(255, 150, 0))
    Shared ActiveSliderBrush2B As Brush = New SolidBrush(Color.Red)
    Shared InactiveSliderBrush1 As Brush = New SolidBrush(Color.FromArgb(70, Color.Gray))
    Shared InactiveSliderBrush2 As Brush = New SolidBrush(Color.FromArgb(50, Color.Black))
    Dim DisplayPadding As Integer = 5
    Dim MarkWidth As Integer = 6
    Const LabelMarkSize As Integer = 7

    Dim minStep As Integer
    Dim maxStep As Integer
    Dim center As Integer
    Dim totalSteps As Integer
    Dim stepSize As Single

    Private Sub CenterTextAt(ByVal gr As Graphics, ByVal txt As _
String, ByVal x As Single, ByVal y As Single)
        ' Mark the center for debugging.
        'gr.DrawLine(Pens.Red, x - 10, y, x + 10, y)
        'gr.DrawLine(Pens.Red, x, y - 10, x, y + 10)

        ' Make a StringFormat object that centers.
        Dim sf As New StringFormat
        sf.LineAlignment = StringAlignment.Center
        sf.Alignment = StringAlignment.Center

        ' Draw the text.
        gr.DrawString(txt, Fonte, Brushes.White, x, y, sf)
        sf.Dispose()
    End Sub

    Private Sub Calculos()

        Try

            minStep = CInt(Math.Truncate(Math.Floor(GetToneStep(MinFrequency))))
            maxStep = CInt(Math.Truncate(Math.Ceiling(GetToneStep(MaxFrequency))))

            center = Width \ 4

            totalSteps = maxStep - minStep
            stepSize = CSng(Me.Height - 2 * DisplayPadding) / totalSteps

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Public Sub GerarImagemDeFundo()

        Try
            MinFrequency = 40 '65 '70
            MaxFrequency = 4000 '1200
            DisplayPadding = 5
            MarkWidth = 6


            Dim FaceBit As New Bitmap(Me.Width, Me.Height)
            Dim gr As Graphics = Graphics.FromImage(FaceBit)


            Calculos()



            Dim Borda As Pen = New Pen(Brushes.White, 3)
            Dim rect As Rectangle = New Rectangle(0, 0, Me.Width, CInt(Me.ClientRectangle.Height / 2))
            Dim lgBrush As LinearGradientBrush = New LinearGradientBrush(rect, Color.FromArgb(220, 50, 50, 50), Color.FromArgb(220, 50, 50, 50), LinearGradientMode.Vertical)
            Dim rect2 As Rectangle = New Rectangle(0, CInt(Me.ClientRectangle.Height / 2), Me.Width, CInt(Me.ClientRectangle.Height / 2) + 10)
            Dim lgBrush2 As LinearGradientBrush = New LinearGradientBrush(rect2, Color.FromArgb(230, 0, 0, 0), Color.FromArgb(255, 0, 0, 0), LinearGradientMode.Vertical)
            gr.FillRectangle(lgBrush, rect)
            gr.FillRectangle(lgBrush2, rect2)

            gr.DrawRectangle(Borda, 0, 1, Me.Width - 3, Me.Height - 3)
            gr.DrawRectangle(Borda, 0, 1, CInt(Me.Width / 2), Me.Height - 3)

            For i As Integer = 0 To totalSteps
                Dim y As Single = stepSize * i + DisplayPadding

                gr.DrawLine(MarkerPen, center - MarkWidth \ 2, y, center + MarkWidth \ 2, y)
            Next


            gr.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit


            Dim maxTextWidth As Single = gr.MeasureString("WW", Fonte).Width
            For Each label As ScaleLabel In Labels
                Dim labelBrush As Brush = New SolidBrush(label.Color)
                Dim labelStep As Double = GetToneStep(label.Frequency)
                Dim labelPosition As Single = CSng(stepSize * (maxStep - labelStep) + DisplayPadding)

                Dim labelMarkPosition As New RectangleF(DisplayPadding, (labelPosition - LabelMarkSize \ 2) - 1, LabelMarkSize, LabelMarkSize)

                If label.Title = "B7" OrElse label.Title = "B5" OrElse label.Title = "B3" OrElse label.Title = "B1" Then
                    gr.FillRectangle(New SolidBrush(Color.FromArgb(35, 255, 255, 255)), CInt((Me.Width / 2)) - 23, (labelPosition - LabelMarkSize \ 2) - (stepSize / 4), 15, stepSize * 12)
                End If

                gr.SmoothingMode = SmoothingMode.AntiAlias
                If label.Color <> Color.Transparent Then
                    gr.FillEllipse(labelBrush, labelMarkPosition)
                    gr.DrawEllipse(MarkerPen, labelMarkPosition)
                    gr.FillEllipse(Brushes.White, DisplayPadding + LabelMarkSize \ 5, labelPosition - LabelMarkSize \ 3, LabelMarkSize \ 3, LabelMarkSize \ 3)
                End If

                Dim titleSize As SizeF = gr.MeasureString(label.Title, Fonte)

                CenterTextAt(gr, label.Title, (Width / 2) - 15, labelPosition + 1)

            Next



            Me.BackgroundImage = FaceBit

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
        'base.OnPaint(e);

        Try

            MinFrequency = 40
            MaxFrequency = 4000
            DisplayPadding = 5

            Calculos()




            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias

            If CurrentFrequency > 0 AndAlso ClosestFrequency > 0 Then

                Dim sliderBrush1 As Brush, sliderBrush2 As Brush

                Dim MargemErroInferior, MargemErroSuperior As Double
                MargemErroInferior = ClosestFrequency / (2 ^ (1 / (((MemoNotes.TrackBar3.Maximum * 2) / (MemoNotes.TrackBar3.Maximum - MemoNotes.TrackBar3.Value)) * 12)))
                MargemErroSuperior = ClosestFrequency * (2 ^ (1 / (((MemoNotes.TrackBar3.Maximum * 2) / (MemoNotes.TrackBar3.Maximum - MemoNotes.TrackBar3.Value)) * 12)))


                If Not SignalDetected Then
                    sliderBrush1 = InactiveSliderBrush1
                    sliderBrush2 = InactiveSliderBrush2
                Else
                    If CurrentFrequency >= MargemErroInferior AndAlso CurrentFrequency <= MargemErroSuperior Then
                        sliderBrush1 = ActiveSliderBrush1
                        sliderBrush2 = ActiveSliderBrush2
                    Else
                        sliderBrush1 = ActiveSliderBrush1B
                        sliderBrush2 = ActiveSliderBrush2B
                    End If
                End If



                Dim sliderStep As Double = GetToneStep(CurrentFrequency)
                Dim sliderPosition As Single = CSng(stepSize * (maxStep - sliderStep) + DisplayPadding)

                e.Graphics.FillPolygon(sliderBrush1, New PointF() {New PointF(center - 10, sliderPosition), New PointF(center, sliderPosition - 5), New PointF(center, sliderPosition + 5), New PointF(center + 10, sliderPosition)})
                e.Graphics.FillPolygon(sliderBrush2, New PointF() {New PointF(center - 10, sliderPosition), New PointF(center, sliderPosition + 5), New PointF(center, sliderPosition - 5), New PointF(center + 10, sliderPosition)})


                If ClosestFrequency.ToString("f3") = 82.407 OrElse ClosestFrequency.ToString("f3") = 110.0 OrElse ClosestFrequency.ToString("f3") = 146.832 OrElse _
                    ClosestFrequency.ToString("f3") = 195.998 OrElse ClosestFrequency.ToString("f3") = 246.942 OrElse ClosestFrequency.ToString("f3") = 329.628 OrElse _
                    ClosestFrequency.ToString("f3") = 261.626 OrElse ClosestFrequency.ToString("f3") = 440.0 Then

                    Dim labelStep As Double = GetToneStep(ClosestFrequency)
                    Dim labelPosition As Single = CSng(stepSize * (maxStep - labelStep) + DisplayPadding)

                    Dim labelMarkPosition As New RectangleF(DisplayPadding, (labelPosition - LabelMarkSize \ 2) - 1, LabelMarkSize, LabelMarkSize)
                    e.Graphics.FillEllipse(Brushes.Yellow, labelMarkPosition)
                    e.Graphics.DrawEllipse(MarkerPen, labelMarkPosition)
                    e.Graphics.FillEllipse(Brushes.White, DisplayPadding + LabelMarkSize \ 5, labelPosition - LabelMarkSize \ 3, LabelMarkSize \ 3, LabelMarkSize \ 3)

                End If





                MinFrequency = ClosestFrequency / (2 ^ (1 / 24))
                MaxFrequency = ClosestFrequency * (2 ^ (1 / 24))
                DisplayPadding = 10

                e.Graphics.SmoothingMode = SmoothingMode.None

                Dim LinhaMargemErro As Pen = New Pen(Brushes.GreenYellow, 1)
                LinhaMargemErro.DashStyle = DashStyle.Dot
                e.Graphics.DrawLine(LinhaMargemErro, CInt(Me.Width / 2) + 2, CInt((Me.ClientSize.Height / 2) - ((Me.ClientSize.Height - (DisplayPadding * 2)) / ((MemoNotes.TrackBar3.Maximum * 2) / (MemoNotes.TrackBar3.Maximum - MemoNotes.TrackBar3.Value)))), Me.Width - 5, CInt((Me.ClientSize.Height / 2) - ((Me.ClientSize.Height - (DisplayPadding * 2)) / ((MemoNotes.TrackBar3.Maximum * 2) / (MemoNotes.TrackBar3.Maximum - MemoNotes.TrackBar3.Value)))))
                e.Graphics.DrawLine(LinhaMargemErro, CInt(Me.Width / 2) + 2, CInt((Me.ClientSize.Height / 2) + ((Me.ClientSize.Height - (DisplayPadding * 2)) / ((MemoNotes.TrackBar3.Maximum * 2) / (MemoNotes.TrackBar3.Maximum - MemoNotes.TrackBar3.Value)))), Me.Width - 5, CInt((Me.ClientSize.Height / 2) + ((Me.ClientSize.Height - (DisplayPadding * 2)) / ((MemoNotes.TrackBar3.Maximum * 2) / (MemoNotes.TrackBar3.Maximum - MemoNotes.TrackBar3.Value)))))
                e.Graphics.DrawString("1/" & FormatNumber(((MemoNotes.TrackBar3.Maximum * 2) / (MemoNotes.TrackBar3.Maximum - MemoNotes.TrackBar3.Value)), 2) & " semitom", New Font("Arial", 6), Brushes.GreenYellow, New PointF(CInt(Me.Width / 2) + 5, CInt((Me.ClientSize.Height / 2) - ((Me.ClientSize.Height - (DisplayPadding * 2)) / ((MemoNotes.TrackBar3.Maximum * 2) / (MemoNotes.TrackBar3.Maximum - MemoNotes.TrackBar3.Value))))))


                Calculos()

                Dim Dimensão As Integer = Me.ClientRectangle.Height - 20
                Dim fonte2 As Font = New Font("Arial", 8, FontStyle.Bold)

                Dim labelPosition2 As Single = CSng(Me.ClientRectangle.Height / 2) - 1
                Dim labelMarkPosition2 As New RectangleF(DisplayPadding + (center * 2), labelPosition2 - LabelMarkSize \ 2, LabelMarkSize, LabelMarkSize)
                Dim titleSize2 As SizeF = e.Graphics.MeasureString(NoteName, fonte2)




                For z As Integer = 0 To 20
                    If z = 0 OrElse z = 10 OrElse z = 20 Then
                        MarkWidth = 10
                        e.Graphics.DrawLine(MarkerPen2, (center * 3) - MarkWidth \ 2, CInt(Dimensão * (z / 20)) + 10, (center * 3) + MarkWidth \ 2, CInt(Dimensão * (z / 20)) + 10)

                    Else
                        MarkWidth = 2
                        e.Graphics.DrawLine(MarkerPen, (center * 3) - MarkWidth \ 2, CInt(Dimensão * (z / 20)) + 10, (center * 3) + MarkWidth \ 2, CInt(Dimensão * (z / 20)) + 10)
                    End If

                    e.Graphics.DrawString((ClosestFrequency * (2 ^ (1 / (12 * (20 / (10 - z)))))).ToString("F3"), New Font("Arial", 6), Brushes.White, New PointF((center * 3) + MarkWidth + 6 \ 2, CInt(Dimensão * (z / 20)) + 6))

                Next

                e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias


                e.Graphics.FillEllipse(Brushes.GreenYellow, labelMarkPosition2)
                e.Graphics.DrawEllipse(MarkerPen, labelMarkPosition2)
                e.Graphics.FillEllipse(Brushes.White, DisplayPadding + (center * 2) + LabelMarkSize \ 5, labelPosition2 - LabelMarkSize \ 3, LabelMarkSize \ 3, LabelMarkSize \ 3)
                e.Graphics.DrawString(NoteName, fonte2, Brushes.SkyBlue, New PointF((center * 3) - 60 \ 2, labelPosition2 + 1 - titleSize2.Height / 2))


                totalSteps = maxStep - minStep
                stepSize = CSng(Me.ClientRectangle.Height - 2 * DisplayPadding) / (totalSteps - 1)
                sliderStep = GetToneStep(CurrentFrequency)
                sliderPosition = CSng(stepSize * (maxStep - sliderStep) + (DisplayPadding + 10 - (Me.ClientRectangle.Height / 2)))

                e.Graphics.FillPolygon(sliderBrush1, New PointF() {New PointF((center * 3) - 10, sliderPosition), New PointF((center * 3), sliderPosition - 5), New PointF((center * 3), sliderPosition + 5), New PointF((center * 3) + 10, sliderPosition)})
                e.Graphics.FillPolygon(sliderBrush2, New PointF() {New PointF((center * 3) - 10, sliderPosition), New PointF((center * 3), sliderPosition + 5), New PointF((center * 3), sliderPosition - 5), New PointF((center * 3) + 10, sliderPosition)})



            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Function GetToneStep(ByVal frequency As Double) As Double
        Return Math.Log(frequency / AFrequency, ToneStep)
    End Function

    Private Class ScaleLabel
        Public Title As String
        Public Frequency As Double
        Public Color As Color
    End Class

    Private Sub FrequenciesScale_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        GerarImagemDeFundo()
    End Sub

End Class