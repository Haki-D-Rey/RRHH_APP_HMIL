<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Frm_Catalogos_Seleccion
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Frm_Catalogos_Seleccion))
        Me.PanelPrincipalA = New System.Windows.Forms.Panel()
        Me.TV = New System.Windows.Forms.TreeView()
        Me.btnPROCESOS = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.PanelPrincipalA.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PanelPrincipalA
        '
        Me.PanelPrincipalA.BackColor = System.Drawing.Color.DarkCyan
        Me.PanelPrincipalA.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PanelPrincipalA.Controls.Add(Me.TV)
        Me.PanelPrincipalA.Controls.Add(Me.btnPROCESOS)
        Me.PanelPrincipalA.Dock = System.Windows.Forms.DockStyle.Left
        Me.PanelPrincipalA.Location = New System.Drawing.Point(0, 0)
        Me.PanelPrincipalA.Name = "PanelPrincipalA"
        Me.PanelPrincipalA.Size = New System.Drawing.Size(342, 785)
        Me.PanelPrincipalA.TabIndex = 79
        Me.PanelPrincipalA.Tag = "9999"
        '
        'TV
        '
        Me.TV.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TV.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TV.Indent = 20
        Me.TV.ItemHeight = 20
        Me.TV.Location = New System.Drawing.Point(3, 38)
        Me.TV.Name = "TV"
        Me.TV.Size = New System.Drawing.Size(334, 740)
        Me.TV.TabIndex = 2
        '
        'btnPROCESOS
        '
        Me.btnPROCESOS.BackColor = System.Drawing.Color.DarkCyan
        Me.btnPROCESOS.Dock = System.Windows.Forms.DockStyle.Top
        Me.btnPROCESOS.FlatAppearance.BorderSize = 0
        Me.btnPROCESOS.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(23, Byte), Integer), CType(CType(21, Byte), Integer), CType(CType(32, Byte), Integer))
        Me.btnPROCESOS.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(24, Byte), Integer), CType(CType(22, Byte), Integer), CType(CType(34, Byte), Integer))
        Me.btnPROCESOS.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnPROCESOS.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPROCESOS.ForeColor = System.Drawing.Color.White
        Me.btnPROCESOS.Image = CType(resources.GetObject("btnPROCESOS.Image"), System.Drawing.Image)
        Me.btnPROCESOS.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnPROCESOS.Location = New System.Drawing.Point(0, 0)
        Me.btnPROCESOS.Name = "btnPROCESOS"
        Me.btnPROCESOS.Padding = New System.Windows.Forms.Padding(5, 0, 0, 0)
        Me.btnPROCESOS.Size = New System.Drawing.Size(340, 32)
        Me.btnPROCESOS.TabIndex = 1
        Me.btnPROCESOS.Tag = "9999"
        Me.btnPROCESOS.Text = "Catalogos Asignados"
        Me.btnPROCESOS.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnPROCESOS.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnPROCESOS.UseVisualStyleBackColor = False
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.BackColor = System.Drawing.Color.White
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(0, 0)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(1319, 824)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox1.TabIndex = 83
        Me.PictureBox1.TabStop = False
        '
        'Frm_Catalogos_Seleccion
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1303, 785)
        Me.Controls.Add(Me.PanelPrincipalA)
        Me.Controls.Add(Me.PictureBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Frm_Catalogos_Seleccion"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Control de Catalogos"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.PanelPrincipalA.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents PanelPrincipalA As Panel
    Private WithEvents btnPROCESOS As Button
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents TV As TreeView
End Class
