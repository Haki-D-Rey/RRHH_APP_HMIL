<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_Visor_Reportes
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
        Me.CrystalR = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.SuspendLayout()
        '
        'CrystalR
        '
        Me.CrystalR.ActiveViewIndex = -1
        Me.CrystalR.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CrystalR.Cursor = System.Windows.Forms.Cursors.Default
        Me.CrystalR.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CrystalR.Location = New System.Drawing.Point(0, 0)
        Me.CrystalR.Name = "CrystalR"
        Me.CrystalR.ShowParameterPanelButton = False
        Me.CrystalR.ShowRefreshButton = False
        Me.CrystalR.Size = New System.Drawing.Size(1077, 660)
        Me.CrystalR.TabIndex = 3
        Me.CrystalR.Tag = "9999"
        Me.CrystalR.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
        '
        'Frm_Visor_Reportes
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1077, 660)
        Me.Controls.Add(Me.CrystalR)
        Me.Name = "Frm_Visor_Reportes"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Reportes"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents CrystalR As CrystalDecisions.Windows.Forms.CrystalReportViewer
End Class
