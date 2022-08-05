<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_Visor_Documentos
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Frm_Visor_Documentos))
        Me.Skin = New Sunisoft.IrisSkin.SkinEngine(CType(Me, System.ComponentModel.Component))
        Me.PDFcontendor = New AxAcroPDFLib.AxAcroPDF()
        CType(Me.PDFcontendor, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Skin
        '
        Me.Skin.SerialNumber = "U4N2UjLguUZs33UR+Vy47JAZ81t2fjIFvut28vc5oHiVeivGb/NZMA=="
        Me.Skin.SkinAllForm = False
        Me.Skin.SkinFile = "D:\Proyectos\SIRH\Componentes Graficos Vb2\SKIN NET 2010 WIN 7\SkinVS.NET\Sports\" &
    "SportsBlue.ssk"
        Me.Skin.SkinStreamMain = CType(resources.GetObject("Skin.SkinStreamMain"), System.IO.Stream)
        '
        'PDFcontendor
        '
        Me.PDFcontendor.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PDFcontendor.Enabled = True
        Me.PDFcontendor.Location = New System.Drawing.Point(0, 0)
        Me.PDFcontendor.Name = "PDFcontendor"
        Me.PDFcontendor.OcxState = CType(resources.GetObject("PDFcontendor.OcxState"), System.Windows.Forms.AxHost.State)
        Me.PDFcontendor.Size = New System.Drawing.Size(770, 703)
        Me.PDFcontendor.TabIndex = 1
        '
        'Frm_Visor_Documentos
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(770, 703)
        Me.Controls.Add(Me.PDFcontendor)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Frm_Visor_Documentos"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Visor de Documentos"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.PDFcontendor, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Skin As Sunisoft.IrisSkin.SkinEngine
    Friend WithEvents PDFcontendor As AxAcroPDFLib.AxAcroPDF
End Class
