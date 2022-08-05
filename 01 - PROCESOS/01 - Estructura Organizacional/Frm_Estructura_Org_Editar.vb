Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Public Class Frm_Estructura_Org_Editar
    Private Sub Frm_Estructura_Org_Editar_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.CheckBox1.Checked = False
            Me.CheckBox2.Checked = False
            Me.CheckBox3.Checked = False
            Me.CheckBox4.Checked = False
            Me.TextBox37.Text = ""
            Me.ComboBox1.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Estructura_Org_Editar_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = vmcsID_M_C
        Me.CARGAR_COMBO_CAT_DE_CARGOS_DE_ESTRUCTURA()
        Me.ComboBox1.Text = vmcsD_CARGO_ES
        Me.CheckBox1.Checked = vmcsJEFE
        Me.CARGAR_COMBO_CAT_DE_GRADO_PLANTILLA()
        Me.ComboBox2.Text = vmcsD_GP
        Me.CARGAR_COMBO_CAT_DE_GRADO_REAL()
        Me.ComboBox3.Text = vmcsD_GR
        Me.CARGAR_COMBO_CAT_DE_CATEGORIA_SALARIAL()
        Me.ComboBox4.Text = vmcsD_CAT_SALARIAL
        Me.CARGAR_COMBO_CAT_DE_SITUACION()
        Me.ComboBox5.Text = vmcsD_SITUACION
        Me.CheckBox2.Checked = vmcsMIXTA
        Me.CheckBox3.Checked = vmcsMILITAR
        Me.CheckBox4.Checked = vmcsPAME
        Me.TextBox37.Text = vmcsN_ORDEN_MILITAR
        Call CARGAR_COMBO_CAT_DE_ESTABLECIMIENTO()
        Me.ComboBox8.Text = vmcsD_ESTABLECIMIENTO
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_ESTRUCTURA()
        Me.ComboBox7.Text = vmcsD_T_ES
        Me.ComboBox1.Focus()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_ESTRUCTURA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE ESTRUCTURA] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox7
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_T_ES"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.CheckBox1.Checked = False
        Me.CheckBox2.Checked = False
        Me.CheckBox3.Checked = False
        Me.CheckBox4.Checked = False
        Me.TextBox37.Text = ""
        Me.ComboBox1.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTABLECIMIENTO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTABLECIMIENTO] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox8
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_ESTABLECIMIENTO"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_CARGOS_DE_ESTRUCTURA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE CARGOS DE ESTRUCTURA] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_CARGO_ES"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_GRADO_PLANTILLA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE GRADO PLANTILLA] order by ID_GP", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox2
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_GP"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_GRADO_REAL()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE GRADO REAL] order by ID_GR", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox3
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_GR"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_CATEGORIA_SALARIAL()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE CATEGORIA SALARIAL] order by ID_CAT_SALARIAL", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox4
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_CAT_SALARIAL"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_SITUACION()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE SITUACION] order by ID_SITUACION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox5
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_SITUACION"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CheckBox3_CheckedChanged(sender As Object, e As System.EventArgs) Handles CheckBox3.CheckedChanged
        Me.TextBox37.Enabled = Me.CheckBox3.Checked
    End Sub
    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub CheckBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox1.KeyDown
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
    Private Sub ComboBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub CheckBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub CheckBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub CheckBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox37_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox37.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox7.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox8.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "0" Then
            MsgBox("No se ha generado el código del Registro", vbInformation, "Mensaje del Sistema")
            Me.Button1.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("El Cargo no ha sido seleccionado Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox2.SelectedValue) = 0 Then
            MsgBox("El Grado Plantilla no ha sido seleccionado Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox2.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox3.SelectedValue) = 0 Then
            MsgBox("El Grado Real no ha sido seleccionado Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox3.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox4.SelectedValue) = 0 Then
            MsgBox("La Categoria Salarial no ha sido seleccionada Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox4.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox5.SelectedValue) = 0 Then
            MsgBox("La Situacion del Cargo no ha sido seleccionada Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox5.Focus()
            Exit Sub
        End If
        If Me.CheckBox3.Checked = False And Me.CheckBox4.Checked = False Then
            MsgBox("Debe Seleccionar el Tipo de Estructura", vbInformation, "Mensaje del Sistema")
            Me.CheckBox4.Focus()
            Exit Sub
        End If
        If Me.CheckBox3.Checked = True And Val(Me.TextBox37.Text) = 0 Then
            MsgBox("Debe Digitar el No. de Orden Militar para el cargo que desea Agregar", vbInformation, "Mensaje del Sistema")
            Me.TextBox37.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox8.SelectedValue) = 0 Then
            MsgBox("El Establecimiento no ha sido seleccionado Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox8.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox7.SelectedValue) = 0 Then
            MsgBox("El Tipo de Estructura del Cargo seleccionada no es Correcta", vbInformation, "Mensaje del Sistema")
            Me.ComboBox7.Focus()
            Exit Sub
        End If
        Call EDITAR_REGISTRO()
        POS = Me.TextBox1.Text
        Me.TextBox1.Text = "0"
        Call CARGAR_DATAGRIDVIEW_3()
        Call ARMAR_DATAGRIDVIEW_3()
        Call SELECCIONAR_DATAGRIDVIEW_3()
        Me.CheckBox1.Checked = False
        Me.CheckBox2.Checked = False
        Me.CheckBox3.Checked = False
        Me.CheckBox4.Checked = False
        Me.TextBox37.Text = ""
        Me.ComboBox1.Focus()
        CERRAR = False
        Me.Close()
    End Sub
    Dim POS As Integer
    Private Sub CARGAR_DATAGRIDVIEW_3()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql3, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE CARGOS SIT]")
        Frm_Estructura_Org.DGV3.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE CARGOS SIT]")
        Frm_Estructura_Org.Label4.Text = "Total de Cargos: " & Frm_Estructura_Org.DGV3.RowCount
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_3()
        Frm_Estructura_Org.DGV3.Columns(0).HeaderText = "Id"
        Frm_Estructura_Org.DGV3.Columns(0).Width = 20
        Frm_Estructura_Org.DGV3.Columns(0).Visible = True
        Frm_Estructura_Org.DGV3.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Estructura_Org.DGV3.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Frm_Estructura_Org.DGV3.Columns(1).HeaderText = "ORDEN"
        Frm_Estructura_Org.DGV3.Columns(1).Width = 0
        Frm_Estructura_Org.DGV3.Columns(1).Visible = False

        Frm_Estructura_Org.DGV3.Columns(2).HeaderText = "ID_N1"
        Frm_Estructura_Org.DGV3.Columns(2).Width = 0
        Frm_Estructura_Org.DGV3.Columns(2).Visible = False

        Frm_Estructura_Org.DGV3.Columns(3).HeaderText = "ORDEN_N1"
        Frm_Estructura_Org.DGV3.Columns(3).Width = 0
        Frm_Estructura_Org.DGV3.Columns(3).Visible = False

        Frm_Estructura_Org.DGV3.Columns(4).HeaderText = "Nivel 1"
        Frm_Estructura_Org.DGV3.Columns(4).Width = 0
        Frm_Estructura_Org.DGV3.Columns(4).Visible = False
        Frm_Estructura_Org.DGV3.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(5).HeaderText = "ID_N2"
        Frm_Estructura_Org.DGV3.Columns(5).Width = 0
        Frm_Estructura_Org.DGV3.Columns(5).Visible = False

        Frm_Estructura_Org.DGV3.Columns(6).HeaderText = "ORDEN_N2"
        Frm_Estructura_Org.DGV3.Columns(6).Width = 0
        Frm_Estructura_Org.DGV3.Columns(6).Visible = False

        Frm_Estructura_Org.DGV3.Columns(7).HeaderText = "Nivel 2"
        Frm_Estructura_Org.DGV3.Columns(7).Width = 0
        Frm_Estructura_Org.DGV3.Columns(7).Visible = False
        Frm_Estructura_Org.DGV3.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(8).HeaderText = "ID_N3"
        Frm_Estructura_Org.DGV3.Columns(8).Width = 0
        Frm_Estructura_Org.DGV3.Columns(8).Visible = False

        Frm_Estructura_Org.DGV3.Columns(9).HeaderText = "ORDEN_N3"
        Frm_Estructura_Org.DGV3.Columns(9).Width = 0
        Frm_Estructura_Org.DGV3.Columns(9).Visible = False

        Frm_Estructura_Org.DGV3.Columns(10).HeaderText = "Nivel 3"
        Frm_Estructura_Org.DGV3.Columns(10).Width = 0
        Frm_Estructura_Org.DGV3.Columns(10).Visible = False
        Frm_Estructura_Org.DGV3.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(11).HeaderText = "ID_N4"
        Frm_Estructura_Org.DGV3.Columns(11).Width = 0
        Frm_Estructura_Org.DGV3.Columns(11).Visible = False

        Frm_Estructura_Org.DGV3.Columns(12).HeaderText = "ORDEN_N4"
        Frm_Estructura_Org.DGV3.Columns(12).Width = 0
        Frm_Estructura_Org.DGV3.Columns(12).Visible = False

        Frm_Estructura_Org.DGV3.Columns(13).HeaderText = "Nivel 4"
        Frm_Estructura_Org.DGV3.Columns(13).Width = 0
        Frm_Estructura_Org.DGV3.Columns(13).Visible = False
        Frm_Estructura_Org.DGV3.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(14).HeaderText = "ID_N5"
        Frm_Estructura_Org.DGV3.Columns(14).Width = 0
        Frm_Estructura_Org.DGV3.Columns(14).Visible = False

        Frm_Estructura_Org.DGV3.Columns(15).HeaderText = "ORDEN_N5"
        Frm_Estructura_Org.DGV3.Columns(15).Width = 0
        Frm_Estructura_Org.DGV3.Columns(15).Visible = False

        Frm_Estructura_Org.DGV3.Columns(16).HeaderText = "Nivel 5"
        Frm_Estructura_Org.DGV3.Columns(16).Width = 0
        Frm_Estructura_Org.DGV3.Columns(16).Visible = False
        Frm_Estructura_Org.DGV3.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(17).HeaderText = "ID_N6"
        Frm_Estructura_Org.DGV3.Columns(17).Width = 0
        Frm_Estructura_Org.DGV3.Columns(17).Visible = False

        Frm_Estructura_Org.DGV3.Columns(18).HeaderText = "ORDEN_N6"
        Frm_Estructura_Org.DGV3.Columns(18).Width = 0
        Frm_Estructura_Org.DGV3.Columns(18).Visible = False

        Frm_Estructura_Org.DGV3.Columns(19).HeaderText = "Nivel 6"
        Frm_Estructura_Org.DGV3.Columns(19).Width = 0
        Frm_Estructura_Org.DGV3.Columns(19).Visible = False
        Frm_Estructura_Org.DGV3.Columns(19).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(20).HeaderText = "ID_N7"
        Frm_Estructura_Org.DGV3.Columns(20).Width = 0
        Frm_Estructura_Org.DGV3.Columns(20).Visible = False

        Frm_Estructura_Org.DGV3.Columns(21).HeaderText = "ORDEN_N7"
        Frm_Estructura_Org.DGV3.Columns(21).Width = 0
        Frm_Estructura_Org.DGV3.Columns(21).Visible = False

        Frm_Estructura_Org.DGV3.Columns(22).HeaderText = "Nivel 7"
        Frm_Estructura_Org.DGV3.Columns(22).Width = 0
        Frm_Estructura_Org.DGV3.Columns(22).Visible = False
        Frm_Estructura_Org.DGV3.Columns(22).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(23).HeaderText = "ID_N8"
        Frm_Estructura_Org.DGV3.Columns(23).Width = 0
        Frm_Estructura_Org.DGV3.Columns(23).Visible = False

        Frm_Estructura_Org.DGV3.Columns(24).HeaderText = "ORDEN_N8"
        Frm_Estructura_Org.DGV3.Columns(24).Width = 0
        Frm_Estructura_Org.DGV3.Columns(24).Visible = False

        Frm_Estructura_Org.DGV3.Columns(25).HeaderText = "Nivel 8"
        Frm_Estructura_Org.DGV3.Columns(25).Width = 0
        Frm_Estructura_Org.DGV3.Columns(25).Visible = False
        Frm_Estructura_Org.DGV3.Columns(25).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(26).HeaderText = "ID_T_ES"
        Frm_Estructura_Org.DGV3.Columns(26).Width = 0
        Frm_Estructura_Org.DGV3.Columns(26).Visible = False

        Frm_Estructura_Org.DGV3.Columns(27).HeaderText = "D_T_ES"
        Frm_Estructura_Org.DGV3.Columns(27).Width = 0
        Frm_Estructura_Org.DGV3.Columns(27).Visible = False

        Frm_Estructura_Org.DGV3.Columns(28).HeaderText = "No. Orden [MIXTO]"
        Frm_Estructura_Org.DGV3.Columns(28).Width = 50
        Frm_Estructura_Org.DGV3.Columns(28).Visible = True
        Frm_Estructura_Org.DGV3.Columns(28).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Estructura_Org.DGV3.Columns(28).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Estructura_Org.DGV3.Columns(28).DefaultCellStyle.BackColor = Color.Khaki

        Frm_Estructura_Org.DGV3.Columns(29).HeaderText = "No. Orden [MILITAR]"
        Frm_Estructura_Org.DGV3.Columns(29).Width = 50
        Frm_Estructura_Org.DGV3.Columns(29).Visible = True
        Frm_Estructura_Org.DGV3.Columns(29).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Estructura_Org.DGV3.Columns(29).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Estructura_Org.DGV3.Columns(29).DefaultCellStyle.BackColor = Color.LightGreen

        Frm_Estructura_Org.DGV3.Columns(30).HeaderText = "No. Orden [PAME]"
        Frm_Estructura_Org.DGV3.Columns(30).Width = 50
        Frm_Estructura_Org.DGV3.Columns(30).Visible = True
        Frm_Estructura_Org.DGV3.Columns(30).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Estructura_Org.DGV3.Columns(30).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Estructura_Org.DGV3.Columns(30).DefaultCellStyle.BackColor = Color.LightBlue

        Frm_Estructura_Org.DGV3.Columns(31).HeaderText = "ID_CARGO_ES"
        Frm_Estructura_Org.DGV3.Columns(31).Width = 0
        Frm_Estructura_Org.DGV3.Columns(31).Visible = False

        Frm_Estructura_Org.DGV3.Columns(32).HeaderText = "Cargo"
        Frm_Estructura_Org.DGV3.Columns(32).Width = 180
        Frm_Estructura_Org.DGV3.Columns(32).Visible = True
        Frm_Estructura_Org.DGV3.Columns(32).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(33).HeaderText = "ID_GP"
        Frm_Estructura_Org.DGV3.Columns(33).Width = 0
        Frm_Estructura_Org.DGV3.Columns(33).Visible = False

        Frm_Estructura_Org.DGV3.Columns(34).HeaderText = "Grado Plantilla"
        Frm_Estructura_Org.DGV3.Columns(34).Width = 80
        Frm_Estructura_Org.DGV3.Columns(34).Visible = True
        Frm_Estructura_Org.DGV3.Columns(34).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(35).HeaderText = "ID_GR"
        Frm_Estructura_Org.DGV3.Columns(35).Width = 0
        Frm_Estructura_Org.DGV3.Columns(35).Visible = False

        Frm_Estructura_Org.DGV3.Columns(36).HeaderText = "Grado Real"
        Frm_Estructura_Org.DGV3.Columns(36).Width = 80
        Frm_Estructura_Org.DGV3.Columns(36).Visible = True
        Frm_Estructura_Org.DGV3.Columns(36).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(37).HeaderText = "ID_CAT_SALARIAL"
        Frm_Estructura_Org.DGV3.Columns(37).Width = 0
        Frm_Estructura_Org.DGV3.Columns(37).Visible = False

        Frm_Estructura_Org.DGV3.Columns(38).HeaderText = "D_CAT_SALARIAL"
        Frm_Estructura_Org.DGV3.Columns(38).Width = 0
        Frm_Estructura_Org.DGV3.Columns(38).Visible = False

        Frm_Estructura_Org.DGV3.Columns(39).HeaderText = "ID_M_P"
        Frm_Estructura_Org.DGV3.Columns(39).Width = 0
        Frm_Estructura_Org.DGV3.Columns(39).Visible = False

        Frm_Estructura_Org.DGV3.Columns(40).HeaderText = "Codigo"
        Frm_Estructura_Org.DGV3.Columns(40).Width = 80
        Frm_Estructura_Org.DGV3.Columns(40).Visible = True
        Frm_Estructura_Org.DGV3.Columns(40).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(41).HeaderText = "Apellidos y Nombres"
        Frm_Estructura_Org.DGV3.Columns(41).Width = 220
        Frm_Estructura_Org.DGV3.Columns(41).Visible = True
        Frm_Estructura_Org.DGV3.Columns(41).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Estructura_Org.DGV3.Columns(41).DefaultCellStyle.BackColor = Color.Khaki

        Frm_Estructura_Org.DGV3.Columns(42).HeaderText = "Nombres y Apellidos"
        Frm_Estructura_Org.DGV3.Columns(42).Width = 0
        Frm_Estructura_Org.DGV3.Columns(42).Visible = False

        Frm_Estructura_Org.DGV3.Columns(43).HeaderText = "ID_SITUACION"
        Frm_Estructura_Org.DGV3.Columns(43).Width = 0
        Frm_Estructura_Org.DGV3.Columns(43).Visible = False

        Frm_Estructura_Org.DGV3.Columns(44).HeaderText = "Situacion"
        Frm_Estructura_Org.DGV3.Columns(44).Width = 60
        Frm_Estructura_Org.DGV3.Columns(44).Visible = True
        Frm_Estructura_Org.DGV3.Columns(44).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(45).HeaderText = "Jefe"
        Frm_Estructura_Org.DGV3.Columns(45).Width = 40
        Frm_Estructura_Org.DGV3.Columns(45).Visible = True
        Frm_Estructura_Org.DGV3.Columns(45).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(46).HeaderText = "Mixta"
        Frm_Estructura_Org.DGV3.Columns(46).Width = 0
        Frm_Estructura_Org.DGV3.Columns(46).Visible = False
        Frm_Estructura_Org.DGV3.Columns(46).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(47).HeaderText = "Militar"
        Frm_Estructura_Org.DGV3.Columns(47).Width = 40
        Frm_Estructura_Org.DGV3.Columns(47).Visible = True
        Frm_Estructura_Org.DGV3.Columns(47).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(48).HeaderText = "Pame"
        Frm_Estructura_Org.DGV3.Columns(48).Width = 40
        Frm_Estructura_Org.DGV3.Columns(48).Visible = True
        Frm_Estructura_Org.DGV3.Columns(48).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(49).HeaderText = "ID_ESTABLECIMIENTO"
        Frm_Estructura_Org.DGV3.Columns(49).Width = 0
        Frm_Estructura_Org.DGV3.Columns(49).Visible = False

        Frm_Estructura_Org.DGV3.Columns(50).HeaderText = "Establecimiento"
        Frm_Estructura_Org.DGV3.Columns(50).Width = 120
        Frm_Estructura_Org.DGV3.Columns(50).Visible = True
        Frm_Estructura_Org.DGV3.Columns(50).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(51).HeaderText = "D_ADMIN"
        Frm_Estructura_Org.DGV3.Columns(51).Width = 0
        Frm_Estructura_Org.DGV3.Columns(51).Visible = False
    End Sub
    Private Sub SELECCIONAR_DATAGRIDVIEW_3()
        Dim I As Integer
        For I = 0 To Frm_Estructura_Org.DGV3.RowCount - 1
            If Frm_Estructura_Org.DGV3.Item(0, I).Value = POS Then
                Frm_Estructura_Org.DGV3.Rows(I).Selected = True
                Frm_Estructura_Org.DGV3.CurrentCell = Frm_Estructura_Org.DGV3.Item(0, I)
                vmcsORDENt = Frm_Estructura_Org.DGV3.Item(1, I).Value
                Exit Sub
            End If
        Next
    End Sub
    Private Sub EDITAR_REGISTRO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE CARGOS] SET ID_T_ES = " & Me.ComboBox7.SelectedValue & ", ID_CARGO_ES = " & Me.ComboBox1.SelectedValue & ", JEFE = '" & Me.CheckBox1.Checked & "', ID_GP = " & Me.ComboBox2.SelectedValue & ", ID_GR = " & Me.ComboBox3.SelectedValue & ", ID_CAT_SALARIAL = " & Me.ComboBox4.SelectedValue & ", ID_SITUACION = " & Me.ComboBox5.SelectedValue & ", MIXTA = '" & Me.CheckBox2.Checked & "', MILITAR = '" & Me.CheckBox3.Checked & "', PAME = '" & Me.CheckBox4.Checked & "', N_ORDEN_MILITAR = '" & Me.TextBox37.Text & "', ID_ESTABLECIMIENTO = " & Me.ComboBox8.SelectedValue & "  WHERE ID_M_C = " & CInt(Me.TextBox1.Text) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
End Class