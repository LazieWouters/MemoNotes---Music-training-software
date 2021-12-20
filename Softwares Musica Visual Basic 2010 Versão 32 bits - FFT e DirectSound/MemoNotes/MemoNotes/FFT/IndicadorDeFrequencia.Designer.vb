<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class IndicadorDeFrequencia
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(IndicadorDeFrequencia))
        Me.FrequenciesScale1 = New Global.MemoNotes.FrequenciesScale(Me.components)
        Me.SuspendLayout()
        '
        'FrequenciesScale1
        '
        Me.FrequenciesScale1.BackColor = System.Drawing.Color.White
        Me.FrequenciesScale1.BackgroundImage = CType(resources.GetObject("FrequenciesScale1.BackgroundImage"), System.Drawing.Image)
        Me.FrequenciesScale1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.FrequenciesScale1.Location = New System.Drawing.Point(1, 0)
        Me.FrequenciesScale1.Name = "FrequenciesScale1"
        Me.FrequenciesScale1.NoteName = Nothing
        Me.FrequenciesScale1.Size = New System.Drawing.Size(278, 755)
        Me.FrequenciesScale1.TabIndex = 0
        '
        'IndicadorDeFrequencia
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = Global.MemoNotes.My.MySettings.Default.TamanhoJanelaIndicadorDeFrequencia
        Me.ControlBox = False
        Me.Controls.Add(Me.FrequenciesScale1)
        Me.DataBindings.Add(New System.Windows.Forms.Binding("Location", Global.MemoNotes.My.MySettings.Default, "PosiçãoJanelaIndicadorDeFrequencia", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.DataBindings.Add(New System.Windows.Forms.Binding("ClientSize", Global.MemoNotes.My.MySettings.Default, "TamanhoJanelaIndicadorDeFrequencia", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.Location = Global.MemoNotes.My.MySettings.Default.PosiçãoJanelaIndicadorDeFrequencia
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "IndicadorDeFrequencia"
        Me.Text = "Indicador De Frequência"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents FrequenciesScale1 As Global.MemoNotes.FrequenciesScale
End Class
