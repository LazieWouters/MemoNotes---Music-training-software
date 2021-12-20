<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TreinamentoDinâmicaMãoEsquerdaDireita
    Inherits Global.MemoNotes.PerPixelAlphaForm

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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(TreinamentoDinâmicaMãoEsquerdaDireita))
        Me.Cms1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.EntradasMIDIToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.Label1 = New System.Windows.Forms.Label
        Me.PPP = New System.Windows.Forms.ColorDialog
        Me.PP = New System.Windows.Forms.ColorDialog
        Me.P = New System.Windows.Forms.ColorDialog
        Me.MP = New System.Windows.Forms.ColorDialog
        Me.MF = New System.Windows.Forms.ColorDialog
        Me.FF = New System.Windows.Forms.ColorDialog
        Me.FFF = New System.Windows.Forms.ColorDialog
        Me.CorPPP = New System.Windows.Forms.PictureBox
        Me.CorPP = New System.Windows.Forms.PictureBox
        Me.CorP = New System.Windows.Forms.PictureBox
        Me.CorMP = New System.Windows.Forms.PictureBox
        Me.CorMF = New System.Windows.Forms.PictureBox
        Me.CorF = New System.Windows.Forms.PictureBox
        Me.CorFF = New System.Windows.Forms.PictureBox
        Me.CorFFF = New System.Windows.Forms.PictureBox
        Me.F = New System.Windows.Forms.ColorDialog
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.CoresPadrão = New System.Windows.Forms.PictureBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Cms1.SuspendLayout()
        CType(Me.CorPPP, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CorPP, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CorP, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CorMP, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CorMF, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CorF, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CorFF, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CorFFF, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CoresPadrão, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
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
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Location = New System.Drawing.Point(20, 395)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 41)
        Me.Label1.TabIndex = 1
        '
        'PPP
        '
        Me.PPP.AnyColor = True
        Me.PPP.Color = Global.MemoNotes.My.MySettings.Default.NovoValorPPP
        Me.PPP.FullOpen = True
        '
        'PP
        '
        Me.PP.AnyColor = True
        Me.PP.Color = Global.MemoNotes.My.MySettings.Default.NovoValorPP
        Me.PP.FullOpen = True
        '
        'P
        '
        Me.P.AnyColor = True
        Me.P.Color = Global.MemoNotes.My.MySettings.Default.NovoValorP
        Me.P.FullOpen = True
        '
        'MP
        '
        Me.MP.AnyColor = True
        Me.MP.Color = Global.MemoNotes.My.MySettings.Default.NovoValorMP
        Me.MP.FullOpen = True
        '
        'MF
        '
        Me.MF.AnyColor = True
        Me.MF.Color = Global.MemoNotes.My.MySettings.Default.NovoValorMF
        Me.MF.FullOpen = True
        Me.MF.ShowHelp = True
        '
        'FF
        '
        Me.FF.AnyColor = True
        Me.FF.Color = Global.MemoNotes.My.MySettings.Default.NovoValorFF
        Me.FF.FullOpen = True
        '
        'FFF
        '
        Me.FFF.AnyColor = True
        Me.FFF.Color = Global.MemoNotes.My.MySettings.Default.NovoValorFFF
        Me.FFF.FullOpen = True
        '
        'CorPPP
        '
        Me.CorPPP.BackColor = System.Drawing.Color.Transparent
        Me.CorPPP.Cursor = System.Windows.Forms.Cursors.Hand
        Me.CorPPP.Location = New System.Drawing.Point(157, 208)
        Me.CorPPP.Name = "CorPPP"
        Me.CorPPP.Size = New System.Drawing.Size(94, 46)
        Me.CorPPP.TabIndex = 690
        Me.CorPPP.TabStop = False
        Me.ToolTip1.SetToolTip(Me.CorPPP, "Duplo clique para alterar a cor")
        '
        'CorPP
        '
        Me.CorPP.BackColor = System.Drawing.Color.Transparent
        Me.CorPP.Cursor = System.Windows.Forms.Cursors.Hand
        Me.CorPP.Location = New System.Drawing.Point(251, 208)
        Me.CorPP.Name = "CorPP"
        Me.CorPP.Size = New System.Drawing.Size(94, 46)
        Me.CorPP.TabIndex = 691
        Me.CorPP.TabStop = False
        Me.ToolTip1.SetToolTip(Me.CorPP, "Duplo clique para alterar a cor")
        '
        'CorP
        '
        Me.CorP.BackColor = System.Drawing.Color.Transparent
        Me.CorP.Cursor = System.Windows.Forms.Cursors.Hand
        Me.CorP.Location = New System.Drawing.Point(345, 208)
        Me.CorP.Name = "CorP"
        Me.CorP.Size = New System.Drawing.Size(94, 46)
        Me.CorP.TabIndex = 692
        Me.CorP.TabStop = False
        Me.ToolTip1.SetToolTip(Me.CorP, "Duplo clique para alterar a cor")
        '
        'CorMP
        '
        Me.CorMP.BackColor = System.Drawing.Color.Transparent
        Me.CorMP.Cursor = System.Windows.Forms.Cursors.Hand
        Me.CorMP.Location = New System.Drawing.Point(438, 208)
        Me.CorMP.Name = "CorMP"
        Me.CorMP.Size = New System.Drawing.Size(94, 46)
        Me.CorMP.TabIndex = 693
        Me.CorMP.TabStop = False
        Me.ToolTip1.SetToolTip(Me.CorMP, "Duplo clique para alterar a cor")
        '
        'CorMF
        '
        Me.CorMF.BackColor = System.Drawing.Color.Transparent
        Me.CorMF.Cursor = System.Windows.Forms.Cursors.Hand
        Me.CorMF.Location = New System.Drawing.Point(532, 208)
        Me.CorMF.Name = "CorMF"
        Me.CorMF.Size = New System.Drawing.Size(94, 46)
        Me.CorMF.TabIndex = 694
        Me.CorMF.TabStop = False
        Me.ToolTip1.SetToolTip(Me.CorMF, "Duplo clique para alterar a cor")
        '
        'CorF
        '
        Me.CorF.BackColor = System.Drawing.Color.Transparent
        Me.CorF.Cursor = System.Windows.Forms.Cursors.Hand
        Me.CorF.Location = New System.Drawing.Point(624, 208)
        Me.CorF.Name = "CorF"
        Me.CorF.Size = New System.Drawing.Size(94, 46)
        Me.CorF.TabIndex = 695
        Me.CorF.TabStop = False
        Me.ToolTip1.SetToolTip(Me.CorF, "Duplo clique para alterar a cor")
        '
        'CorFF
        '
        Me.CorFF.BackColor = System.Drawing.Color.Transparent
        Me.CorFF.Cursor = System.Windows.Forms.Cursors.Hand
        Me.CorFF.Location = New System.Drawing.Point(717, 208)
        Me.CorFF.Name = "CorFF"
        Me.CorFF.Size = New System.Drawing.Size(94, 46)
        Me.CorFF.TabIndex = 696
        Me.CorFF.TabStop = False
        Me.ToolTip1.SetToolTip(Me.CorFF, "Duplo clique para alterar a cor")
        '
        'CorFFF
        '
        Me.CorFFF.BackColor = System.Drawing.Color.Transparent
        Me.CorFFF.Cursor = System.Windows.Forms.Cursors.Hand
        Me.CorFFF.Location = New System.Drawing.Point(811, 208)
        Me.CorFFF.Name = "CorFFF"
        Me.CorFFF.Size = New System.Drawing.Size(94, 46)
        Me.CorFFF.TabIndex = 697
        Me.CorFFF.TabStop = False
        Me.ToolTip1.SetToolTip(Me.CorFFF, "Duplo clique para alterar a cor")
        '
        'F
        '
        Me.F.AnyColor = True
        Me.F.Color = Global.MemoNotes.My.MySettings.Default.NovoValorF
        Me.F.FullOpen = True
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
        'CoresPadrão
        '
        Me.CoresPadrão.BackColor = System.Drawing.Color.Transparent
        Me.CoresPadrão.Cursor = System.Windows.Forms.Cursors.Hand
        Me.CoresPadrão.Location = New System.Drawing.Point(483, 259)
        Me.CoresPadrão.Name = "CoresPadrão"
        Me.CoresPadrão.Size = New System.Drawing.Size(100, 13)
        Me.CoresPadrão.TabIndex = 698
        Me.CoresPadrão.TabStop = False
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Location = New System.Drawing.Point(957, 33)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(34, 18)
        Me.Label2.TabIndex = 700
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Location = New System.Drawing.Point(919, 33)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(34, 18)
        Me.Label3.TabIndex = 699
        '
        'TreinamentoDinâmicaMãoEsquerdaDireita
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.MemoNotes.My.Resources.Resources.TreinamentoDinamica
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.ClientSize = New System.Drawing.Size(1062, 555)
        Me.ContextMenuStrip = Me.Cms1
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.CorFF)
        Me.Controls.Add(Me.CoresPadrão)
        Me.Controls.Add(Me.CorFFF)
        Me.Controls.Add(Me.CorF)
        Me.Controls.Add(Me.CorMF)
        Me.Controls.Add(Me.CorMP)
        Me.Controls.Add(Me.CorP)
        Me.Controls.Add(Me.CorPP)
        Me.Controls.Add(Me.CorPPP)
        Me.Controls.Add(Me.Label1)
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "TreinamentoDinâmicaMãoEsquerdaDireita"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Treinamento Dinâmica Mão Esquerda Direita"
        Me.Cms1.ResumeLayout(False)
        CType(Me.CorPPP, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CorPP, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CorP, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CorMP, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CorMF, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CorF, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CorFF, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CorFFF, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CoresPadrão, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Cms1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents EntradasMIDIToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents PPP As System.Windows.Forms.ColorDialog
    Friend WithEvents PP As System.Windows.Forms.ColorDialog
    Friend WithEvents P As System.Windows.Forms.ColorDialog
    Friend WithEvents MP As System.Windows.Forms.ColorDialog
    Friend WithEvents MF As System.Windows.Forms.ColorDialog
    Friend WithEvents FF As System.Windows.Forms.ColorDialog
    Friend WithEvents FFF As System.Windows.Forms.ColorDialog
    Friend WithEvents CorPPP As System.Windows.Forms.PictureBox
    Friend WithEvents CorPP As System.Windows.Forms.PictureBox
    Friend WithEvents CorP As System.Windows.Forms.PictureBox
    Friend WithEvents CorMP As System.Windows.Forms.PictureBox
    Friend WithEvents CorMF As System.Windows.Forms.PictureBox
    Friend WithEvents CorF As System.Windows.Forms.PictureBox
    Friend WithEvents CorFF As System.Windows.Forms.PictureBox
    Friend WithEvents CorFFF As System.Windows.Forms.PictureBox
    Friend WithEvents F As System.Windows.Forms.ColorDialog
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents CoresPadrão As System.Windows.Forms.PictureBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
End Class
