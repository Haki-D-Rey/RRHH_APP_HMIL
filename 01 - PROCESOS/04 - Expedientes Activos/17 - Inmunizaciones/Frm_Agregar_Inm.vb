﻿Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Agregar_Inm
    Private Sub Frm_Agregar_Inm_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = "0"
            Me.ComboBox1.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Agregar_Inm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = "0"
        Call CARGAR_COMBO_CAT_DE_VACUNAS()
        Call CARGAR_COMBO_CAT_DE_DOSIS()
        Call Clases.BUSCAR_FECHA_Y_HORA()
        Me.DateTimePicker1.Value = Mid(FECHA_SERVIDOR, 1, 10)
        Me.DateTimePicker1.Checked = False
        Me.ComboBox1.Focus()
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Me.TextBox1.Text = "0"
        Me.ComboBox1.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_DOSIS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE DOSIS] order by ID_DOSIS", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox2
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_DOSIS"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_VACUNAS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE VACUNAS] order by ID_VACUNA", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_VACUNA"
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
    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub DateTimePicker1_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker1.KeyDown
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
            MiDataAdapter = New SqlDataAdapter("SELECT ID_M_INM FROM [MAESTRO DE INMUNIZACIONES] ORDER BY ID_M_INM", Conexion.CxRRHH)
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
            Me.Button7.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar la Vacuna Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox2.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar la Dosis Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox2.Focus()
            Exit Sub
        End If
        Call GRABAR_REGISTRO()
        inID_M_INM = Me.TextBox1.Text
        CADENAsql = "SELECT * FROM [dbo].[VISTA MAESTRO DE INMUNIZACIONES] WHERE ID_M_P = " & vID_M_P & " ORDER BY [ID_M_INM]"
        Call CARGAR_DATAGRIDVIEW_19()
        Call ARMAR_DATAGRIDVIEW_19()
        Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_19()
        Me.TextBox1.Text = "0"
        Me.ComboBox1.Focus()
    End Sub
    Private Sub GRABAR_REGISTRO()
        Dim fFECHA_I As String
        If Me.DateTimePicker1.Checked = True Then
            fFECHA_I = Mid(Me.DateTimePicker1.Value, 1, 10)
        Else
            fFECHA_I = "NULL"
        End If
        Dim f1 As Object = If(fFECHA_I <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_I + "', 105), 23)", "NULL")


        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()

        Dim SQL As String
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            SQL = "INSERT INTO [MAESTRO DE INMUNIZACIONES] (ID_M_INM, ID_M_P, ID_VACUNA, ID_DOSIS, FECHA_V)" &
                   "values (" & Me.TextBox1.Text & ", " & vID_M_P & ", " & Me.ComboBox1.SelectedValue & ", " & Me.ComboBox2.SelectedValue & ", " & f1 & ")"
            Dim cmd As New SqlCommand(SQL, CxRRHH)
            cmd.ExecuteNonQuery()
            cmd = Nothing

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_19()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE INMUNIZACIONES]")
        Frm_Expedientes_Activos.DGV19.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE INMUNIZACIONES]")
        Frm_Expedientes_Activos.Label116.Text = Frm_Expedientes_Activos.DGV19.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_19()
        Frm_Expedientes_Activos.DGV19.Columns(0).HeaderText = "Id"
        Frm_Expedientes_Activos.DGV19.Columns(0).Width = 30
        Frm_Expedientes_Activos.DGV19.Columns(0).Visible = True
        Frm_Expedientes_Activos.DGV19.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV19.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Frm_Expedientes_Activos.DGV19.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Frm_Expedientes_Activos.DGV19.Columns(1).HeaderText = "ID_M_P"
        Frm_Expedientes_Activos.DGV19.Columns(1).Width = 0
        Frm_Expedientes_Activos.DGV19.Columns(1).Visible = False

        Frm_Expedientes_Activos.DGV19.Columns(2).HeaderText = "ID_VACUNA"
        Frm_Expedientes_Activos.DGV19.Columns(2).Width = 0
        Frm_Expedientes_Activos.DGV19.Columns(2).Visible = False

        Frm_Expedientes_Activos.DGV19.Columns(3).HeaderText = "Vacuna"
        Frm_Expedientes_Activos.DGV19.Columns(3).Width = 200
        Frm_Expedientes_Activos.DGV19.Columns(3).Visible = True
        Frm_Expedientes_Activos.DGV19.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV19.Columns(4).HeaderText = "ID_DOSIS"
        Frm_Expedientes_Activos.DGV19.Columns(4).Width = 0
        Frm_Expedientes_Activos.DGV19.Columns(4).Visible = False

        Frm_Expedientes_Activos.DGV19.Columns(5).HeaderText = "Dosis"
        Frm_Expedientes_Activos.DGV19.Columns(5).Width = 80
        Frm_Expedientes_Activos.DGV19.Columns(5).Visible = True
        Frm_Expedientes_Activos.DGV19.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV19.Columns(6).HeaderText = "Fecha Vacuna"
        Frm_Expedientes_Activos.DGV19.Columns(6).Width = 80
        Frm_Expedientes_Activos.DGV19.Columns(6).Visible = True
        Frm_Expedientes_Activos.DGV19.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV19.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_19()
        Dim I As Integer
        For I = 0 To Frm_Expedientes_Activos.DGV19.RowCount - 1
            If Frm_Expedientes_Activos.DGV19.Item(0, I).Value = inID_M_INM Then
                Frm_Expedientes_Activos.DGV19.Rows(I).Selected = True
                Frm_Expedientes_Activos.DGV19.CurrentCell = Frm_Expedientes_Activos.DGV19.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
End Class