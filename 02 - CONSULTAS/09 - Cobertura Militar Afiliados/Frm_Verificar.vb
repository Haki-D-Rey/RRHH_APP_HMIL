Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Verificar
    Private Sub Frm_Verificar_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            COBERTURA_AFILIADOS = False
            Me.TextBox1.SelectAll()
            Me.TextBox1.Focus()
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Verificar_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Cursor = Cursors.WaitCursor
        Application.DoEvents()
        COBERTURA_AFILIADOS = True
        RemoveHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
        RemoveHandler Me.ComboBox2.SelectedIndexChanged, AddressOf ComboBox2_SelectedIndexChanged
        Call CARGAR_COMBO_SIE_CAT_DE_BUSQUEDA_TIPO()
        Call CARGAR_COMBO_SIE_CAT_DE_BUSQUEDA()
        Me.ComboBox2.Text = "NOMBRES Y APELLIDOS"
        Me.TextBox1.Text = "DIGITE EL TEXTO A BUSCAR..."
        Me.DGV.DataSource = Nothing
        AddHandler Me.ComboBox2.SelectedIndexChanged, AddressOf ComboBox2_SelectedIndexChanged
        AddHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
        Me.TextBox1.SelectAll()
        Me.TextBox1.Focus()
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW()
        If Conexion.CxSIE.State = ConnectionState.Open Then
            Conexion.CBD_SIE()
        End If
        Conexion.ABD_SIE()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxSIE)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[SIE VISTA MAESTRO DE CONSULTA PRINCIPAL]")
        DGV.DataSource = MiDataSet.Tables("[SIE VISTA MAESTRO DE CONSULTA PRINCIPAL]")
        If Conexion.CxSIE.State = ConnectionState.Open Then
            Conexion.CBD_SIE()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW()
        Me.DGV.Columns(0).HeaderText = "Tipo de Cobertura"
        Me.DGV.Columns(0).Width = 150
        Me.DGV.Columns(0).Visible = True
        Me.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(1).HeaderText = "Tipo de Ingreso"
        Me.DGV.Columns(1).Width = 100
        Me.DGV.Columns(1).Visible = True
        Me.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(2).HeaderText = "Nss"
        Me.DGV.Columns(2).Width = 60
        Me.DGV.Columns(2).Visible = True
        Me.DGV.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(3).HeaderText = "Nombres"
        Me.DGV.Columns(3).Width = 0
        Me.DGV.Columns(3).Visible = False
        Me.DGV.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(4).HeaderText = "Apellidos"
        Me.DGV.Columns(4).Width = 0
        Me.DGV.Columns(4).Visible = False
        Me.DGV.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(5).HeaderText = "Nombres y Apellidos"
        Me.DGV.Columns(5).Width = 300
        Me.DGV.Columns(5).Visible = True
        Me.DGV.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(6).HeaderText = ""
        Me.DGV.Columns(6).Width = 10
        Me.DGV.Columns(6).Visible = True
        Me.DGV.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(6).DefaultCellStyle.BackColor = Color.Blue
        Me.DGV.Columns(6).Frozen = True

        Me.DGV.Columns(7).HeaderText = "Identidad (8 Digitos)"
        Me.DGV.Columns(7).Width = 80
        Me.DGV.Columns(7).Visible = True
        Me.DGV.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(8).HeaderText = "Sexo"
        Me.DGV.Columns(8).Width = 40
        Me.DGV.Columns(8).Visible = True
        Me.DGV.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(9).HeaderText = "Grado Militar"
        Me.DGV.Columns(9).Width = 150
        Me.DGV.Columns(9).Visible = True
        Me.DGV.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(10).HeaderText = "Tipo de Parentela"
        Me.DGV.Columns(10).Width = 80
        Me.DGV.Columns(10).Visible = True
        Me.DGV.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(11).HeaderText = "Unidad Militar"
        Me.DGV.Columns(11).Width = 60
        Me.DGV.Columns(11).Visible = True
        Me.DGV.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(12).HeaderText = "Cargo Militar"
        Me.DGV.Columns(12).Width = 150
        Me.DGV.Columns(12).Visible = True
        Me.DGV.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(13).HeaderText = "Identidad (13 Digitos)"
        Me.DGV.Columns(13).Width = 100
        Me.DGV.Columns(13).Visible = True
        Me.DGV.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(14).HeaderText = "Clasificacion"
        Me.DGV.Columns(14).Width = 80
        Me.DGV.Columns(14).Visible = True
        Me.DGV.Columns(14).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(14).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(15).HeaderText = "No. Cedula"
        Me.DGV.Columns(15).Width = 120
        Me.DGV.Columns(15).Visible = True
        Me.DGV.Columns(15).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(15).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(16).HeaderText = "No. Carne Beneficiario"
        Me.DGV.Columns(16).Width = 80
        Me.DGV.Columns(16).Visible = True
        Me.DGV.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(16).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(17).HeaderText = "Fecha de Ingreso"
        Me.DGV.Columns(17).Width = 75
        Me.DGV.Columns(17).Visible = True
        Me.DGV.Columns(17).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(17).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(18).HeaderText = "Edad"
        Me.DGV.Columns(18).Width = 40
        Me.DGV.Columns(18).Visible = True
        Me.DGV.Columns(18).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(18).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(19).HeaderText = "Fecha de Nacimiento"
        Me.DGV.Columns(19).Width = 75
        Me.DGV.Columns(19).Visible = True
        Me.DGV.Columns(19).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(19).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Me.TextBox1.SelectAll()
        Me.TextBox1.Focus()
        COBERTURA_AFILIADOS = False
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_SIE_CAT_DE_BUSQUEDA_TIPO()
        If Conexion.CxSIE.State = ConnectionState.Open Then
            Conexion.CBD_SIE()
        End If
        Conexion.ABD_SIE()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT DISTINCT TIPO FROM [dbo].[SIE CAT DE BUSQUEDA] WHERE TIPO = 'COBERTURA MILITAR' OR TIPO = 'INSS' ORDER BY TIPO", Conexion.CxSIE)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "TIPO"
                .ValueMember = "TIPO"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxSIE.State = ConnectionState.Open Then
                Conexion.CBD_SIE()
            End If
        End Try
    End Sub
    Dim CADENA_COMBO As String
    Private Sub CARGAR_COMBO_SIE_CAT_DE_BUSQUEDA()
        If Me.ComboBox1.Text = "INSS" Then
            CADENA_COMBO = "SELECT * FROM [SIE CAT DE BUSQUEDA] WHERE TIPO = 'INSS' OR TIPO = 'MIXTO' ORDER BY ORDEN"
        End If
        If Me.ComboBox1.Text = "COBERTURA MILITAR" Then
            CADENA_COMBO = "SELECT * FROM [SIE CAT DE BUSQUEDA] WHERE TIPO = 'COBERTURA MILITAR' OR TIPO = 'MIXTO' ORDER BY ORDEN"
        End If
        'If Me.ComboBox1.Text = "PROGRAMAS ESPECIALES" Then
        '    CADENA_COMBO = "SELECT * FROM [SIE CAT DE BUSQUEDA] WHERE TIPO = 'PROGRAMAS ESPECIALES' OR TIPO = 'MIXTO' ORDER BY ORDEN"
        'End If

        If Conexion.CxSIE.State = ConnectionState.Open Then
            Conexion.CBD_SIE()
        End If
        Conexion.ABD_SIE()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter(CADENA_COMBO, Conexion.CxSIE)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox2
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_BUSQUEDA"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxSIE.State = ConnectionState.Open Then
                Conexion.CBD_SIE()
            End If
        End Try
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        RemoveHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
        RemoveHandler Me.ComboBox2.SelectedIndexChanged, AddressOf ComboBox2_SelectedIndexChanged

        AddHandler Me.ComboBox2.SelectedIndexChanged, AddressOf ComboBox2_SelectedIndexChanged
        AddHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If Len(Me.TextBox1.Text) <> 0 Then
            Me.Cursor = Cursors.WaitCursor
            Application.DoEvents()
            RemoveHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
            RemoveHandler Me.ComboBox2.SelectedIndexChanged, AddressOf ComboBox2_SelectedIndexChanged
            If ComboBox1.Text = "COBERTURA MILITAR" Then
                CADENAsql = "Select * from [SIE VISTA MAESTRO DE CONSULTA PRINCIPAL] WHERE (" & Me.ComboBox2.SelectedValue & " LIKE '%" & Me.TextBox1.Text & "%') AND TIPO = 'COBERTURA MILITAR' ORDER BY TIPO, IDENTIDAD, NOMBRES, APELLIDOS"
            End If
            If ComboBox1.Text = "INSS" Then
                CADENAsql = "Select * from [SIE VISTA MAESTRO DE CONSULTA PRINCIPAL] WHERE (" & Me.ComboBox2.SelectedValue & " LIKE '%" & Me.TextBox1.Text & "%') AND TIPO <> 'COBERTURA MILITAR' ORDER BY TIPO, IDENTIDAD, NOMBRES, APELLIDOS"
            End If
            'If Me.ComboBox1.Text = "PROGRAMAS ESPECIALES" Then
            '    CADENAsql = "Select * from [SIE VISTA MAESTRO DE CONSULTA PRINCIPAL] WHERE (" & Me.ComboBox2.SelectedValue & " LIKE '%" & Me.TextBox1.Text & "%') AND TIPO <> 'PROGRAMAS ESPECIALES' ORDER BY TIPO, IDENTIDAD, NOMBRES, APELLIDOS"
            'End If
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            AddHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
            AddHandler Me.ComboBox2.SelectedIndexChanged, AddressOf ComboBox2_SelectedIndexChanged
            Me.TextBox1.SelectAll()
            Me.TextBox1.Focus()
            Me.Cursor = Cursors.Default
        Else
            MsgBox("Debe ingresar el texto que desea buscar", vbInformation, "Mensaje del Sistema")
            Me.TextBox1.SelectAll()
            Me.DGV.Focus()
        End If
    End Sub
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If (e.KeyCode = Keys.Enter) Then
            If Len(Me.TextBox1.Text) <> 0 Then
                Me.Cursor = Cursors.WaitCursor
                Application.DoEvents()
                RemoveHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
                RemoveHandler Me.ComboBox2.SelectedIndexChanged, AddressOf ComboBox2_SelectedIndexChanged
                If ComboBox1.Text = "COBERTURA MILITAR" Then
                    CADENAsql = "Select * from [SIE VISTA MAESTRO DE CONSULTA PRINCIPAL] WHERE (" & Me.ComboBox2.SelectedValue & " LIKE '%" & Me.TextBox1.Text & "%') AND TIPO = 'COBERTURA MILITAR' ORDER BY TIPO, IDENTIDAD, NOMBRES, APELLIDOS"
                End If
                If ComboBox1.Text = "INSS" Then
                    CADENAsql = "Select * from [SIE VISTA MAESTRO DE CONSULTA PRINCIPAL] WHERE (" & Me.ComboBox2.SelectedValue & " LIKE '%" & Me.TextBox1.Text & "%') AND (TIPO = 'ASEGURADO' OR TIPO = 'RIESGO PROFESIONAL' OR TIPO = 'PENSIONADOS') ORDER BY TIPO, IDENTIDAD, NOMBRES, APELLIDOS"
                End If
                'If Me.ComboBox1.Text = "PROGRAMAS ESPECIALES" Then
                '    CADENAsql = "Select * from [SIE VISTA MAESTRO DE CONSULTA PRINCIPAL] WHERE (" & Me.ComboBox2.SelectedValue & " LIKE '%" & Me.TextBox1.Text & "%') AND TIPO = 'PROGRAMAS ESPECIALES' ORDER BY TIPO, IDENTIDAD, NOMBRES, APELLIDOS"
                'End If
                Call CARGAR_DATAGRIDVIEW()
                Call ARMAR_DATAGRIDVIEW()
                AddHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
                AddHandler Me.ComboBox2.SelectedIndexChanged, AddressOf ComboBox2_SelectedIndexChanged
                Me.TextBox1.SelectAll()
                Me.DGV.Focus()
                Me.Cursor = Cursors.Default
            Else
                MsgBox("Debe ingresar el texto que desea buscar", vbInformation, "Mensaje del Sistema")
                Me.TextBox1.SelectAll()
                Me.TextBox1.Focus()
            End If
        End If
    End Sub
    Private Sub DGV_Click(sender As Object, e As EventArgs) Handles DGV.Click
        If Me.DGV.RowCount > 0 Then
            Dim FILA = DGV.CurrentRow.Index
            vTIPO = Me.DGV.Item(0, FILA).Value
            vTIPO_INGRESO = Me.DGV.Item(1, FILA).Value
            vNSS = Me.DGV.Item(2, FILA).Value
            vNOMBRES = Me.DGV.Item(3, FILA).Value
            vAPELLIDOS = Me.DGV.Item(4, FILA).Value
            vNOMBRES_Y_APELLIDOS = Me.DGV.Item(5, FILA).Value
            vIDENTIDAD = Me.DGV.Item(7, FILA).Value
            vSEXO = Me.DGV.Item(8, FILA).Value
            vGRADO = Me.DGV.Item(9, FILA).Value
            vTIPO_PARENTELA = Me.DGV.Item(10, FILA).Value
            vUM = Me.DGV.Item(11, FILA).Value
            vCARGO = Me.DGV.Item(12, FILA).Value
            vNO_EXPEDIENTE_DPC = Me.DGV.Item(13, FILA).Value
            vCLASIFICACION = Me.DGV.Item(14, FILA).Value
            vCEDULA = Me.DGV.Item(15, FILA).Value
            vNO_CARNE = Me.DGV.Item(16, FILA).Value
            vEDAD = Me.DGV.Item(18, FILA).Value
        End If
    End Sub
    Private Sub DGV_RowPrePaint(sender As Object, e As DataGridViewRowPrePaintEventArgs) Handles DGV.RowPrePaint
        Call FORMATO_FILAS_DATAGRIDVIEW()
    End Sub
    Private Sub FORMATO_FILAS_DATAGRIDVIEW()
        Dim style As New DataGridViewCellStyle
        style.Font = New Font(DGV.Font, FontStyle.Bold)
        For Each Row As DataGridViewRow In DGV.Rows
            If Row.Cells(0).Value = "COBERTURA MILITAR" Then
                Row.DefaultCellStyle.ForeColor = Color.DarkGreen
            End If
            If Row.Cells(0).Value = "ASEGURADO" Then
                Row.DefaultCellStyle.ForeColor = Color.Blue
            End If
            If Row.Cells(0).Value = "RIESGO PROFESIONAL" Then
                Row.DefaultCellStyle.ForeColor = Color.Brown 
            End If
            If Row.Cells(0).Value = "PENSIONADOS" Then
                Row.DefaultCellStyle.ForeColor = Color.DarkRed
            End If
            If Row.Cells(0).Value = "PROGRAMAS ESPECIALES" Then
                Row.DefaultCellStyle.ForeColor = Color.DarkOliveGreen
            End If
        Next
    End Sub
    Private Sub VisualziarTitularYBeneficiariosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VisualziarTitularYBeneficiariosToolStripMenuItem.Click
        If vIDENTIDAD <> "" Then
            Frm_Verificar_Visualizar.ShowDialog()
        Else
            MsgBox("Debe seleccionar el registro del que desea visualizar la Informacion", vbInformation, "Mensaje del Sistema")
            Me.DGV.Focus()
        End If
    End Sub
    Private Sub MovAtlaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopiarNssToolStripMenuItem.Click
        Clipboard.Clear()
        Clipboard.SetText(vNSS)
    End Sub
    Private Sub CopiarNoCédulaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopiarNoCédulaToolStripMenuItem.Click
        Clipboard.Clear()
        Clipboard.SetText(vCEDULA)
    End Sub
    Private Sub CopiarMNoDeIdentidadToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopiarMNoDeIdentidadToolStripMenuItem.Click
        Clipboard.Clear()
        Clipboard.SetText("M" & vIDENTIDAD)
    End Sub
    Private Sub DGV_SelectionChanged(sender As Object, e As EventArgs) Handles DGV.SelectionChanged
        If DGV.SelectedCells.Count <> 0 Then
            Dim FILA = DGV.CurrentRow.Index
            vTIPO = Me.DGV.Item(0, FILA).Value
            vTIPO_INGRESO = Me.DGV.Item(1, FILA).Value
            vNSS = Me.DGV.Item(2, FILA).Value
            vNOMBRES = Me.DGV.Item(3, FILA).Value
            vAPELLIDOS = Me.DGV.Item(4, FILA).Value
            vNOMBRES_Y_APELLIDOS = Me.DGV.Item(5, FILA).Value
            vIDENTIDAD = Me.DGV.Item(7, FILA).Value
            vSEXO = Me.DGV.Item(8, FILA).Value
            vGRADO = Me.DGV.Item(9, FILA).Value
            vTIPO_PARENTELA = Me.DGV.Item(10, FILA).Value
            vUM = Me.DGV.Item(11, FILA).Value
            vCARGO = Me.DGV.Item(12, FILA).Value
            vNO_EXPEDIENTE_DPC = Me.DGV.Item(13, FILA).Value
            vCLASIFICACION = Me.DGV.Item(14, FILA).Value
            vCEDULA = Me.DGV.Item(15, FILA).Value
            vNO_CARNE = Me.DGV.Item(16, FILA).Value
            vEDAD = Me.DGV.Item(18, FILA).Value
        End If
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        RemoveHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
        RemoveHandler Me.ComboBox2.SelectedIndexChanged, AddressOf ComboBox2_SelectedIndexChanged
        Call CARGAR_COMBO_SIE_CAT_DE_BUSQUEDA()
        If Me.ComboBox1.Text = "COBERTURA MILITAR" Then
            Me.ComboBox2.Text = "NOMBRES Y APELLIDOS"
        End If
        If Me.ComboBox1.Text = "INSS" Then
            Me.ComboBox2.Text = "NSS"
        End If
        'If Me.ComboBox1.Text = "PROGRAMAS ESPECIALES" Then
        '    Me.ComboBox2.Text = "NOMBRES Y APELLIDOS"
        'End If
        AddHandler Me.ComboBox2.SelectedIndexChanged, AddressOf ComboBox2_SelectedIndexChanged
        AddHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
    End Sub
    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            Me.TextBox1.Focus()
            Me.TextBox1.SelectAll()
        End If
    End Sub
    Private Sub CopiarNombresYApellidosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopiarNombresYApellidosToolStripMenuItem.Click
        Clipboard.Clear()
        Clipboard.SetText(vNOMBRES_Y_APELLIDOS)
    End Sub
    Private Sub Frm_Verificar_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        COBERTURA_AFILIADOS = False
    End Sub
    Private Sub Frm_Verificar_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        COBERTURA_AFILIADOS = False
    End Sub
    Private Sub Frm_Verificar_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        COBERTURA_AFILIADOS = False
    End Sub
    Private Sub Frm_Verificar_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        COBERTURA_AFILIADOS = False
    End Sub
End Class