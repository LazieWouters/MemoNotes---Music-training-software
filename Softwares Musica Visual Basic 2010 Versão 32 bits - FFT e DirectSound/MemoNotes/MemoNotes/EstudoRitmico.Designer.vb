<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EstudoRitmico
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EstudoRitmico))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Cms1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.EntradasMIDIToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.NumericUpDown1 = New System.Windows.Forms.NumericUpDown()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Cms1.SuspendLayout()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Location = New System.Drawing.Point(112, 728)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(136, 71)
        Me.Label1.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.Label1, "Duplo clique para alterar")
        '
        'ToolTip1
        '
        Me.ToolTip1.AutomaticDelay = 1500
        Me.ToolTip1.AutoPopDelay = 5000
        Me.ToolTip1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(1, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.ToolTip1.InitialDelay = 1500
        Me.ToolTip1.IsBalloon = True
        Me.ToolTip1.ReshowDelay = 1500
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Location = New System.Drawing.Point(252, 728)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(174, 71)
        Me.Label2.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.Label2, "Duplo clique para alterar")
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Location = New System.Drawing.Point(429, 728)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(155, 71)
        Me.Label3.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.Label3, "Duplo clique para alterar")
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.Transparent
        Me.Label6.Location = New System.Drawing.Point(590, 728)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(155, 71)
        Me.Label6.TabIndex = 703
        Me.ToolTip1.SetToolTip(Me.Label6, "Duplo clique para alterar")
        '
        'Cms1
        '
        Me.Cms1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.EntradasMIDIToolStripMenuItem})
        Me.Cms1.Name = "Cms1"
        Me.Cms1.ShowCheckMargin = True
        Me.Cms1.ShowImageMargin = False
        Me.Cms1.ShowItemToolTips = False
        Me.Cms1.Size = New System.Drawing.Size(148, 26)
        '
        'EntradasMIDIToolStripMenuItem
        '
        Me.EntradasMIDIToolStripMenuItem.Enabled = False
        Me.EntradasMIDIToolStripMenuItem.Name = "EntradasMIDIToolStripMenuItem"
        Me.EntradasMIDIToolStripMenuItem.Size = New System.Drawing.Size(147, 22)
        Me.EntradasMIDIToolStripMenuItem.Text = "Entradas MIDI"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(763, 735)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(171, 26)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "Novo exercício"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(938, 735)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(171, 26)
        Me.Button2.TabIndex = 4
        Me.Button2.Text = "Escutar ritmo do exercício"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(938, 767)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(171, 26)
        Me.Button3.TabIndex = 6
        Me.Button3.Text = "Escutar ritmo que você executou"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Location = New System.Drawing.Point(996, 33)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(34, 18)
        Me.Label4.TabIndex = 702
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Location = New System.Drawing.Point(958, 33)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(34, 18)
        Me.Label5.TabIndex = 701
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(593, 757)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(102, 21)
        Me.ProgressBar1.TabIndex = 704
        '
        'NumericUpDown1
        '
        Me.NumericUpDown1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.NumericUpDown1.DataBindings.Add(New System.Windows.Forms.Binding("Value", Global.MemoNotes.My.MySettings.Default, "NovoValorTempoInicial", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.NumericUpDown1.Location = New System.Drawing.Point(223, 761)
        Me.NumericUpDown1.Maximum = New Decimal(New Integer() {300, 0, 0, 0})
        Me.NumericUpDown1.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.NumericUpDown1.Name = "NumericUpDown1"
        Me.NumericUpDown1.Size = New System.Drawing.Size(15, 16)
        Me.NumericUpDown1.TabIndex = 5
        Me.NumericUpDown1.Value = Global.MemoNotes.My.MySettings.Default.NovoValorTempoInicial
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(763, 767)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(171, 26)
        Me.Button4.TabIndex = 705
        Me.Button4.Text = "Opções"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'EstudoRitmico
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.MemoNotes.My.Resources.Resources.TreinamentoRitmico
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.ClientSize = New System.Drawing.Size(1230, 807)
        Me.ContextMenuStrip = Me.Cms1
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.NumericUpDown1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "EstudoRitmico"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Estudo Rítmico"
        Me.Cms1.ResumeLayout(False)
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Cms1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents EntradasMIDIToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents NumericUpDown1 As System.Windows.Forms.NumericUpDown
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents Button4 As System.Windows.Forms.Button
End Class
