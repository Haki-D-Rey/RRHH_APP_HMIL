<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_Verificar
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
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

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Frm_Verificar))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.DGV = New System.Windows.Forms.DataGridView()
        Me.CC = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.VisualziarTitularYBeneficiariosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripSeparator()
        Me.CopiarNssToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CopiarNoCédulaToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CopiarNombresYApellidosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem2 = New System.Windows.Forms.ToolStripSeparator()
        Me.CopiarMNoDeIdentidadToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.Skin = New Sunisoft.IrisSkin.SkinEngine(CType(Me, System.ComponentModel.Component))
        CType(Me.DGV, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.CC.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button7
        '
        Me.Button7.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button7.Image = CType(resources.GetObject("Button7.Image"), System.Drawing.Image)
        Me.Button7.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Button7.Location = New System.Drawing.Point(1114, 4)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(66, 56)
        Me.Button7.TabIndex = 3
        Me.Button7.Text = "&Cerrar"
        Me.Button7.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.Button7.UseVisualStyleBackColor = True
        '
        'DGV
        '
        Me.DGV.AllowUserToAddRows = False
        Me.DGV.AllowUserToDeleteRows = False
        Me.DGV.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DGV.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DGV.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGV.ContextMenuStrip = Me.CC
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DGV.DefaultCellStyle = DataGridViewCellStyle2
        Me.DGV.Location = New System.Drawing.Point(7, 64)
        Me.DGV.MultiSelect = False
        Me.DGV.Name = "DGV"
        Me.DGV.ReadOnly = True
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DGV.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.DGV.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.DGV.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.DGV.Size = New System.Drawing.Size(1173, 565)
        Me.DGV.TabIndex = 2
        '
        'CC
        '
        Me.CC.AllowMerge = False
        Me.CC.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.VisualziarTitularYBeneficiariosToolStripMenuItem, Me.ToolStripMenuItem1, Me.CopiarNssToolStripMenuItem, Me.CopiarNoCédulaToolStripMenuItem, Me.CopiarNombresYApellidosToolStripMenuItem, Me.ToolStripMenuItem2, Me.CopiarMNoDeIdentidadToolStripMenuItem})
        Me.CC.Name = "CMS"
        Me.CC.Size = New System.Drawing.Size(234, 126)
        '
        'VisualziarTitularYBeneficiariosToolStripMenuItem
        '
        Me.VisualziarTitularYBeneficiariosToolStripMenuItem.Name = "VisualziarTitularYBeneficiariosToolStripMenuItem"
        Me.VisualziarTitularYBeneficiariosToolStripMenuItem.Size = New System.Drawing.Size(233, 22)
        Me.VisualziarTitularYBeneficiariosToolStripMenuItem.Text = "Visualizar Información"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(230, 6)
        '
        'CopiarNssToolStripMenuItem
        '
        Me.CopiarNssToolStripMenuItem.Name = "CopiarNssToolStripMenuItem"
        Me.CopiarNssToolStripMenuItem.Size = New System.Drawing.Size(233, 22)
        Me.CopiarNssToolStripMenuItem.Text = "Copiar Nss"
        '
        'CopiarNoCédulaToolStripMenuItem
        '
        Me.CopiarNoCédulaToolStripMenuItem.Name = "CopiarNoCédulaToolStripMenuItem"
        Me.CopiarNoCédulaToolStripMenuItem.Size = New System.Drawing.Size(233, 22)
        Me.CopiarNoCédulaToolStripMenuItem.Text = "Copiar No. Cédula"
        '
        'CopiarNombresYApellidosToolStripMenuItem
        '
        Me.CopiarNombresYApellidosToolStripMenuItem.Name = "CopiarNombresYApellidosToolStripMenuItem"
        Me.CopiarNombresYApellidosToolStripMenuItem.Size = New System.Drawing.Size(233, 22)
        Me.CopiarNombresYApellidosToolStripMenuItem.Text = "Copiar Nombres y Apellidos"
        '
        'ToolStripMenuItem2
        '
        Me.ToolStripMenuItem2.Name = "ToolStripMenuItem2"
        Me.ToolStripMenuItem2.Size = New System.Drawing.Size(230, 6)
        '
        'CopiarMNoDeIdentidadToolStripMenuItem
        '
        Me.CopiarMNoDeIdentidadToolStripMenuItem.Name = "CopiarMNoDeIdentidadToolStripMenuItem"
        Me.CopiarMNoDeIdentidadToolStripMenuItem.Size = New System.Drawing.Size(233, 22)
        Me.CopiarMNoDeIdentidadToolStripMenuItem.Text = "Copiar [M] + No. de Identidad"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(7, 1)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(783, 59)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(5, 1)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(148, 13)
        Me.Label8.TabIndex = 0
        Me.Label8.Text = "Parámetros de Búsqueda"
        '
        'ComboBox1
        '
        Me.ComboBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(12, 33)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(161, 21)
        Me.ComboBox1.TabIndex = 5
        '
        'Button4
        '
        Me.Button4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button4.ForeColor = System.Drawing.Color.Blue
        Me.Button4.Location = New System.Drawing.Point(751, 31)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(33, 24)
        Me.Button4.TabIndex = 1
        Me.Button4.Text = ">"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox1.Location = New System.Drawing.Point(390, 33)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(360, 20)
        Me.TextBox1.TabIndex = 0
        '
        'ComboBox2
        '
        Me.ComboBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Location = New System.Drawing.Point(174, 33)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(214, 21)
        Me.ComboBox2.TabIndex = 6
        '
        'Skin
        '
        Me.Skin.SerialNumber = "U4N2UjLguUZs33UR+Vy47JAZ81t2fjIFvut28vc5oHiVeivGb/NZMA=="
        Me.Skin.SkinFile = "D:\Proyectos\HMSIE\Componentes Graficos Vb2\SKIN NET 2010 WIN 7\SkinVS.NET\Sports" &
    "\SportsCyan.ssk"
        Me.Skin.SkinStreamMain = CType(resources.GetObject("Skin.SkinStreamMain"), System.IO.Stream)
        '
        'Frm_Verificar
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1186, 636)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.ComboBox2)
        Me.Controls.Add(Me.DGV)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.KeyPreview = True
        Me.Name = "Frm_Verificar"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Cobertura Militar\ Afiliados"
        CType(Me.DGV, System.ComponentModel.ISupportInitialize).EndInit()
        Me.CC.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Button7 As Button
    Friend WithEvents DGV As DataGridView
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents ComboBox2 As ComboBox
    Friend WithEvents Label8 As Label
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Button4 As Button
    Friend WithEvents CC As ContextMenuStrip
    Friend WithEvents CopiarNssToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents VisualziarTitularYBeneficiariosToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As ToolStripSeparator
    Friend WithEvents CopiarNoCédulaToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem2 As ToolStripSeparator
    Friend WithEvents CopiarMNoDeIdentidadToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents CopiarNombresYApellidosToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents Skin As Sunisoft.IrisSkin.SkinEngine
End Class
