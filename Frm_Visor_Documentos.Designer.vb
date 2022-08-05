<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Frm_Visor_Documentos
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Frm_Visor_Documentos))
        Me.PDFcontendor = New AxAcroPDFLib.AxAcroPDF()
        CType(Me.PDFcontendor, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PDFcontendor
        '
        Me.PDFcontendor.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PDFcontendor.Enabled = True
        Me.PDFcontendor.Location = New System.Drawing.Point(0, 0)
        Me.PDFcontendor.Name = "PDFcontendor"
        Me.PDFcontendor.OcxState = CType(resources.GetObject("PDFcontendor.OcxState"), System.Windows.Forms.AxHost.State)
        Me.PDFcontendor.Size = New System.Drawing.Size(770, 703)
        Me.PDFcontendor.TabIndex = 0
        '
        'Frm_Visor_Documentos
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(770, 703)
        Me.Controls.Add(Me.PDFcontendor)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.KeyPreview = True
        Me.Name = "Frm_Visor_Documentos"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "Visor de Documentos"
        CType(Me.PDFcontendor, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents PDFcontendor As AxAcroPDFLib.AxAcroPDF
End Class
