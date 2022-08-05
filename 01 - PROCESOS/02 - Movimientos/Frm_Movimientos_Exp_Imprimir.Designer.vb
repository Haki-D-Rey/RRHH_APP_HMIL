<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_Movimientos_Exp_Imprimir
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Frm_Movimientos_Exp_Imprimir))
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.Button03 = New System.Windows.Forms.Button()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Button01 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'CheckBox1
        '
        Me.CheckBox1.Location = New System.Drawing.Point(6, 6)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(265, 20)
        Me.CheckBox1.TabIndex = 13
        Me.CheckBox1.Text = "Imprimir sin Selección de Tipo de Movimientos"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Button03
        '
        Me.Button03.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button03.Image = CType(resources.GetObject("Button03.Image"), System.Drawing.Image)
        Me.Button03.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Button03.Location = New System.Drawing.Point(209, 29)
        Me.Button03.Name = "Button03"
        Me.Button03.Size = New System.Drawing.Size(66, 56)
        Me.Button03.TabIndex = 11
        Me.Button03.Text = "&Cerrar"
        Me.Button03.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.Button03.UseVisualStyleBackColor = True
        '
        'ComboBox1
        '
        Me.ComboBox1.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append
        Me.ComboBox1.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.ComboBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(6, 63)
        Me.ComboBox1.MaxDropDownItems = 20
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(136, 21)
        Me.ComboBox1.TabIndex = 9
        '
        'Label20
        '
        Me.Label20.ForeColor = System.Drawing.Color.Black
        Me.Label20.Location = New System.Drawing.Point(6, 44)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(96, 15)
        Me.Label20.TabIndex = 12
        Me.Label20.Text = "Formato de Salida:"
        '
        'Button01
        '
        Me.Button01.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button01.Image = CType(resources.GetObject("Button01.Image"), System.Drawing.Image)
        Me.Button01.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Button01.Location = New System.Drawing.Point(144, 29)
        Me.Button01.Name = "Button01"
        Me.Button01.Size = New System.Drawing.Size(66, 56)
        Me.Button01.TabIndex = 10
        Me.Button01.Text = "Imprimir"
        Me.Button01.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.Button01.UseVisualStyleBackColor = True
        '
        'Frm_Movimientos_Exp_Imprimir
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 90)
        Me.ControlBox = False
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Button03)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Label20)
        Me.Controls.Add(Me.Button01)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Frm_Movimientos_Exp_Imprimir"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Imprimir"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents Button03 As Button
    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents Label20 As Label
    Friend WithEvents Button01 As Button
End Class
