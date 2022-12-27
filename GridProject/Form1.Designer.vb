<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.dgvPrincipal = New System.Windows.Forms.DataGridView()
        Me.lblC1 = New System.Windows.Forms.Label()
        Me.lblC2 = New System.Windows.Forms.Label()
        Me.lblFecha = New System.Windows.Forms.Label()
        Me.btnImportar = New System.Windows.Forms.Button()
        Me.btnExportar = New System.Windows.Forms.Button()
        Me.menuContextual = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.CortarToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CopiarToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PegarToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.lblC3 = New System.Windows.Forms.Label()
        CType(Me.dgvPrincipal, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.menuContextual.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgvPrincipal
        '
        Me.dgvPrincipal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvPrincipal.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvPrincipal.ContextMenuStrip = Me.menuContextual
        Me.dgvPrincipal.Location = New System.Drawing.Point(13, 13)
        Me.dgvPrincipal.Name = "dgvPrincipal"
        Me.dgvPrincipal.RowHeadersVisible = False
        Me.dgvPrincipal.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvPrincipal.Size = New System.Drawing.Size(535, 314)
        Me.dgvPrincipal.TabIndex = 0
        '
        'lblC1
        '
        Me.lblC1.AutoSize = True
        Me.lblC1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblC1.Location = New System.Drawing.Point(14, 358)
        Me.lblC1.Name = "lblC1"
        Me.lblC1.Size = New System.Drawing.Size(33, 24)
        Me.lblC1.TabIndex = 1
        Me.lblC1.Text = "C1"
        Me.lblC1.Visible = False
        '
        'lblC2
        '
        Me.lblC2.AutoSize = True
        Me.lblC2.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblC2.Location = New System.Drawing.Point(76, 358)
        Me.lblC2.Name = "lblC2"
        Me.lblC2.Size = New System.Drawing.Size(39, 25)
        Me.lblC2.TabIndex = 2
        Me.lblC2.Text = "C2"
        Me.lblC2.Visible = False
        '
        'lblFecha
        '
        Me.lblFecha.AutoSize = True
        Me.lblFecha.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFecha.Location = New System.Drawing.Point(253, 358)
        Me.lblFecha.Name = "lblFecha"
        Me.lblFecha.Size = New System.Drawing.Size(72, 25)
        Me.lblFecha.TabIndex = 3
        Me.lblFecha.Text = "Fecha"
        Me.lblFecha.Visible = False
        '
        'btnImportar
        '
        Me.btnImportar.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnImportar.Location = New System.Drawing.Point(19, 435)
        Me.btnImportar.Name = "btnImportar"
        Me.btnImportar.Size = New System.Drawing.Size(114, 31)
        Me.btnImportar.TabIndex = 4
        Me.btnImportar.Text = "Importar Excel"
        Me.btnImportar.UseVisualStyleBackColor = True
        '
        'btnExportar
        '
        Me.btnExportar.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExportar.Location = New System.Drawing.Point(402, 435)
        Me.btnExportar.Name = "btnExportar"
        Me.btnExportar.Size = New System.Drawing.Size(114, 31)
        Me.btnExportar.TabIndex = 5
        Me.btnExportar.Text = "Exportar a Excel"
        Me.btnExportar.UseVisualStyleBackColor = True
        '
        'menuContextual
        '
        Me.menuContextual.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CortarToolStripMenuItem, Me.CopiarToolStripMenuItem, Me.PegarToolStripMenuItem})
        Me.menuContextual.Name = "menuContextual"
        Me.menuContextual.Size = New System.Drawing.Size(110, 70)
        '
        'CortarToolStripMenuItem
        '
        Me.CortarToolStripMenuItem.Name = "CortarToolStripMenuItem"
        Me.CortarToolStripMenuItem.Size = New System.Drawing.Size(109, 22)
        Me.CortarToolStripMenuItem.Text = "Cortar"
        '
        'CopiarToolStripMenuItem
        '
        Me.CopiarToolStripMenuItem.Name = "CopiarToolStripMenuItem"
        Me.CopiarToolStripMenuItem.Size = New System.Drawing.Size(180, 22)
        Me.CopiarToolStripMenuItem.Text = "Copiar"
        '
        'PegarToolStripMenuItem
        '
        Me.PegarToolStripMenuItem.Name = "PegarToolStripMenuItem"
        Me.PegarToolStripMenuItem.Size = New System.Drawing.Size(180, 22)
        Me.PegarToolStripMenuItem.Text = "Pegar"
        '
        'lblC3
        '
        Me.lblC3.AutoSize = True
        Me.lblC3.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblC3.Location = New System.Drawing.Point(157, 357)
        Me.lblC3.Name = "lblC3"
        Me.lblC3.Size = New System.Drawing.Size(39, 25)
        Me.lblC3.TabIndex = 9
        Me.lblC3.Text = "C3"
        Me.lblC3.Visible = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(560, 481)
        Me.Controls.Add(Me.lblC3)
        Me.Controls.Add(Me.btnExportar)
        Me.Controls.Add(Me.btnImportar)
        Me.Controls.Add(Me.lblFecha)
        Me.Controls.Add(Me.lblC2)
        Me.Controls.Add(Me.lblC1)
        Me.Controls.Add(Me.dgvPrincipal)
        Me.Name = "Form1"
        Me.Text = "Form1"
        CType(Me.dgvPrincipal, System.ComponentModel.ISupportInitialize).EndInit()
        Me.menuContextual.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents dgvPrincipal As DataGridView
    Friend WithEvents lblC1 As Label
    Friend WithEvents lblC2 As Label
    Friend WithEvents lblFecha As Label
    Friend WithEvents btnImportar As Button
    Friend WithEvents btnExportar As Button
    Friend WithEvents menuContextual As ContextMenuStrip
    Friend WithEvents CortarToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents CopiarToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents PegarToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents lblC3 As Label
End Class
