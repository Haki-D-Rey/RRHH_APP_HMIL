Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Agregar_Clasificacion
    Private Sub Frm_Agregar_Clasificacion_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = ""
            Me.TextBox2.Text = ""
            Me.ComboBox1.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Agregar_Clasificacion_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = ""
        Call CARGAR_COMBO_CAT_DE_CLASIFICACION_DE_PERSONAL()
        Me.TextBox2.Text = ""
        Me.ComboBox1.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.ComboBox1.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_CLASIFICACION_DE_PERSONAL()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE CLASIFICACION DE PERSONAL] order by ID_C_PE", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_C_PE"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub GENERAR_CODIGO()
        CODIGO = 0
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            MiDataAdapter = New SqlDataAdapter("SELECT ID_C_PERSONAL FROM [MAESTRO DE CLASIFICACION DE PERSONAL] ORDER BY ID_C_PERSONAL", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                For I = 0 To MiDataTable.Rows.Count - 1
                    CODIGO = MiDataTable.Rows(I).Item(0).ToString
                Next
                CODIGO += 1
            Else
                CODIGO = 1
            End If
            Me.TextBox1.Text = CODIGO
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.TextBox1.Text = "0"
        Call GENERAR_CODIGO()
        If TextBox1.Text = "0" Or TextBox1.Text = "" Then
            MsgBox("No se ha generado el Codigo del registro", vbInformation, "Mensaje del Sistema")
            Me.TextBox1.Text = "0"
            Me.Button2.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Tipo de Clasificación Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox2.Text) = 0 Then
            Me.TextBox2.Text = "."
        End If
        Call GRABAR_REGISTRO()
        vmcoID_C_PERSONAL = Me.TextBox1.Text
        CADENAsql = "SELECT * FROM [VISTA MAESTRO DE CLASIFICACION DE PERSONAL] WHERE ID_M_P = " & Frm_Expedientes_Activos.TextBox35.Text & " ORDER BY ID_M_P, ID_C_PERSONAL"
        Call CARGAR_DATAGRIDVIEW_8()
        Call ARMAR_DATAGRIDVIEW_8()
        Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_8()
        CERRAR = False
    End Sub
    Private Sub GRABAR_REGISTRO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE CLASIFICACION DE PERSONAL] (ID_C_PERSONAL, ID_M_P, ID_C_PE, OBSERVACIONES)" &
                                  "values (" & CInt(Me.TextBox1.Text) & ", " & Frm_Expedientes_Activos.TextBox35.Text & ", " & Me.ComboBox1.SelectedValue & ", '" & Me.TextBox2.Text & "')"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_8()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE CLASIFICACION DE PERSONAL]")
        Frm_Expedientes_Activos.DGV8.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE CLASIFICACION DE PERSONAL]")
        Frm_Expedientes_Activos.Label86.Text = Frm_Expedientes_Activos.DGV8.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_8()
        Frm_Expedientes_Activos.DGV8.Columns(0).HeaderText = "Id"
        Frm_Expedientes_Activos.DGV8.Columns(0).Width = 30
        Frm_Expedientes_Activos.DGV8.Columns(0).Visible = True
        Frm_Expedientes_Activos.DGV8.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV8.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Frm_Expedientes_Activos.DGV8.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Frm_Expedientes_Activos.DGV8.Columns(1).HeaderText = "ID_M_P"
        Frm_Expedientes_Activos.DGV8.Columns(1).Width = 0
        Frm_Expedientes_Activos.DGV8.Columns(1).Visible = False

        Frm_Expedientes_Activos.DGV8.Columns(2).HeaderText = "ID_C_PE"
        Frm_Expedientes_Activos.DGV8.Columns(2).Width = 0
        Frm_Expedientes_Activos.DGV8.Columns(2).Visible = False

        Frm_Expedientes_Activos.DGV8.Columns(3).HeaderText = "Clasificacion"
        Frm_Expedientes_Activos.DGV8.Columns(3).Width = 210
        Frm_Expedientes_Activos.DGV8.Columns(3).Visible = True
        Frm_Expedientes_Activos.DGV8.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV8.Columns(4).HeaderText = "Observaciones"
        Frm_Expedientes_Activos.DGV8.Columns(4).Width = 870
        Frm_Expedientes_Activos.DGV8.Columns(4).Visible = True
        Frm_Expedientes_Activos.DGV8.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_8()
        Dim I As Integer
        For I = 0 To Frm_Expedientes_Activos.DGV8.RowCount - 1
            If Frm_Expedientes_Activos.DGV8.Item(0, I).Value = vmcoID_C_PERSONAL Then
                Frm_Expedientes_Activos.DGV8.Rows(I).Selected = True
                Frm_Expedientes_Activos.DGV8.CurrentCell = Frm_Expedientes_Activos.DGV8.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
End Class