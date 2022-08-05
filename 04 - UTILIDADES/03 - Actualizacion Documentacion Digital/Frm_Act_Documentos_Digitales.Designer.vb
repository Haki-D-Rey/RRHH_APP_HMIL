<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_Act_Documentos_Digitales
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Frm_Act_Documentos_Digitales))
        Me.txtFiltro = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnExaminar = New System.Windows.Forms.Button()
        Me.txtDir = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.toolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnAbrirDir = New System.Windows.Forms.Button()
        Me.btnAbrirFic = New System.Windows.Forms.Button()
        Me.chkConSubDir = New System.Windows.Forms.CheckBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.lvFics = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.chkIgnorarError = New System.Windows.Forms.CheckBox()
        Me.LabelInfo = New System.Windows.Forms.Label()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Skin = New Sunisoft.IrisSkin.SkinEngine(CType(Me, System.ComponentModel.Component))
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'txtFiltro
        '
        Me.txtFiltro.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFiltro.Location = New System.Drawing.Point(118, 11)
        Me.txtFiltro.Name = "txtFiltro"
        Me.txtFiltro.Size = New System.Drawing.Size(608, 20)
        Me.txtFiltro.TabIndex = 13
        Me.txtFiltro.Text = "*.*"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(12, 14)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 16)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Filtro búsqueda:"
        Me.toolTip1.SetToolTip(Me.Label2, "Filtro o especificación para la búsqueda (por ej. *.txt)")
        '
        'btnExaminar
        '
        Me.btnExaminar.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExaminar.Location = New System.Drawing.Point(732, 35)
        Me.btnExaminar.Name = "btnExaminar"
        Me.btnExaminar.Size = New System.Drawing.Size(75, 23)
        Me.btnExaminar.TabIndex = 16
        Me.btnExaminar.Text = "Examinar..."
        Me.toolTip1.SetToolTip(Me.btnExaminar, "Seleccionar el directorio en el que se empezará la búsqueda")
        Me.btnExaminar.UseVisualStyleBackColor = True
        '
        'txtDir
        '
        Me.txtDir.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDir.Location = New System.Drawing.Point(118, 37)
        Me.txtDir.Name = "txtDir"
        Me.txtDir.Size = New System.Drawing.Size(608, 20)
        Me.txtDir.TabIndex = 15
        Me.txtDir.Text = "C:\"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(12, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 16)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "Directorio:"
        Me.toolTip1.SetToolTip(Me.Label1, "Directorio en el que se iniciará la búsqueda")
        '
        'btnAbrirDir
        '
        Me.btnAbrirDir.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnAbrirDir.Location = New System.Drawing.Point(101, 456)
        Me.btnAbrirDir.Name = "btnAbrirDir"
        Me.btnAbrirDir.Size = New System.Drawing.Size(90, 23)
        Me.btnAbrirDir.TabIndex = 22
        Me.btnAbrirDir.Text = "Abrir directorio"
        Me.toolTip1.SetToolTip(Me.btnAbrirDir, "Abrir la carpeta del fichero seleccionado")
        Me.btnAbrirDir.UseVisualStyleBackColor = True
        '
        'btnAbrirFic
        '
        Me.btnAbrirFic.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnAbrirFic.Location = New System.Drawing.Point(12, 456)
        Me.btnAbrirFic.Name = "btnAbrirFic"
        Me.btnAbrirFic.Size = New System.Drawing.Size(90, 23)
        Me.btnAbrirFic.TabIndex = 21
        Me.btnAbrirFic.Text = "Abrir fichero"
        Me.toolTip1.SetToolTip(Me.btnAbrirFic, "Abrir el fichero seleccionado con el programa que tenga asociado")
        Me.btnAbrirFic.UseVisualStyleBackColor = True
        '
        'chkConSubDir
        '
        Me.chkConSubDir.Checked = True
        Me.chkConSubDir.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkConSubDir.Location = New System.Drawing.Point(118, 68)
        Me.chkConSubDir.Name = "chkConSubDir"
        Me.chkConSubDir.Size = New System.Drawing.Size(186, 17)
        Me.chkConSubDir.TabIndex = 17
        Me.chkConSubDir.Text = "Incluir subdirectorios"
        Me.toolTip1.SetToolTip(Me.chkConSubDir, "Si se debe buscar también en los directorios que haya en el directorio indicado")
        Me.chkConSubDir.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(336, 456)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(120, 23)
        Me.Button1.TabIndex = 24
        Me.Button1.Text = "Modificar Nombre"
        Me.toolTip1.SetToolTip(Me.Button1, "Abrir la carpeta del fichero seleccionado")
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Button2.Location = New System.Drawing.Point(455, 456)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(120, 23)
        Me.Button2.TabIndex = 25
        Me.Button2.Text = "Ingresar Documento"
        Me.toolTip1.SetToolTip(Me.Button2, "Abrir la carpeta del fichero seleccionado")
        Me.Button2.UseVisualStyleBackColor = True
        '
        'lvFics
        '
        Me.lvFics.Activation = System.Windows.Forms.ItemActivation.OneClick
        Me.lvFics.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvFics.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2})
        Me.lvFics.FullRowSelect = True
        Me.lvFics.GridLines = True
        Me.lvFics.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.lvFics.HideSelection = False
        Me.lvFics.Location = New System.Drawing.Point(12, 93)
        Me.lvFics.MultiSelect = False
        Me.lvFics.Name = "lvFics"
        Me.lvFics.Size = New System.Drawing.Size(795, 357)
        Me.lvFics.TabIndex = 20
        Me.lvFics.UseCompatibleStateImageBehavior = False
        Me.lvFics.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Nombre"
        Me.ColumnHeader1.Width = 210
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Directorio"
        Me.ColumnHeader2.Width = 180
        '
        'btnBuscar
        '
        Me.btnBuscar.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnBuscar.Location = New System.Drawing.Point(732, 64)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(75, 23)
        Me.btnBuscar.TabIndex = 19
        Me.btnBuscar.Text = "Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = True
        '
        'chkIgnorarError
        '
        Me.chkIgnorarError.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkIgnorarError.AutoSize = True
        Me.chkIgnorarError.Checked = True
        Me.chkIgnorarError.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkIgnorarError.Location = New System.Drawing.Point(579, 68)
        Me.chkIgnorarError.Name = "chkIgnorarError"
        Me.chkIgnorarError.Size = New System.Drawing.Size(147, 17)
        Me.chkIgnorarError.TabIndex = 18
        Me.chkIgnorarError.Text = "Ignorar los avisos de error"
        Me.chkIgnorarError.UseVisualStyleBackColor = True
        Me.chkIgnorarError.Visible = False
        '
        'LabelInfo
        '
        Me.LabelInfo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelInfo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LabelInfo.Location = New System.Drawing.Point(12, 485)
        Me.LabelInfo.Margin = New System.Windows.Forms.Padding(3)
        Me.LabelInfo.Name = "LabelInfo"
        Me.LabelInfo.Size = New System.Drawing.Size(795, 20)
        Me.LabelInfo.TabIndex = 23
        Me.LabelInfo.Text = "LableInfo"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'Skin
        '
        Me.Skin.SerialNumber = "U4N2UjLguUZs33UR+Vy47JAZ81t2fjIFvut28vc5oHiVeivGb/NZMA=="
        Me.Skin.SkinAllForm = False
        Me.Skin.SkinFile = "D:\Proyectos\SPYC\Componentes Graficos Vb2\SKIN NET 2010 WIN 7\SkinVS.NET\Sports\" &
    "SportsCyan.ssk"
        Me.Skin.SkinStreamMain = CType(resources.GetObject("Skin.SkinStreamMain"), System.IO.Stream)
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(575, 457)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(232, 21)
        Me.ComboBox1.TabIndex = 26
        '
        'Frm_Act_Documentos_Digitales
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(819, 517)
        Me.ControlBox = False
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.txtFiltro)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnExaminar)
        Me.Controls.Add(Me.txtDir)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnAbrirDir)
        Me.Controls.Add(Me.btnAbrirFic)
        Me.Controls.Add(Me.lvFics)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.chkConSubDir)
        Me.Controls.Add(Me.chkIgnorarError)
        Me.Controls.Add(Me.LabelInfo)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Frm_Act_Documentos_Digitales"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Actualizar - Documentos Digitales"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents txtFiltro As TextBox
    Private WithEvents Label2 As Label
    Private WithEvents toolTip1 As ToolTip
    Private WithEvents btnExaminar As Button
    Private WithEvents txtDir As TextBox
    Private WithEvents Label1 As Label
    Private WithEvents btnAbrirDir As Button
    Private WithEvents btnAbrirFic As Button
    Private WithEvents chkConSubDir As CheckBox
    Private WithEvents lvFics As ListView
    Private WithEvents ColumnHeader1 As ColumnHeader
    Private WithEvents ColumnHeader2 As ColumnHeader
    Private WithEvents btnBuscar As Button
    Private WithEvents chkIgnorarError As CheckBox
    Private WithEvents LabelInfo As Label
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents Skin As Sunisoft.IrisSkin.SkinEngine
    Private WithEvents Button1 As Button
    Private WithEvents Button2 As Button
    Friend WithEvents ComboBox1 As ComboBox
End Class
