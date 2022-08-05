Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Movimientos_Exp_Add
    Private Sub Frm_Movimientos_Exp_Add_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Movimientos_Exp_Add_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Frm_Movimientos_Exp.ComboBox1.Text = "ALTAS" Then
            Me.Text = "Agregar Registro <ALTAS>"
            Me.ComboBox8.Enabled = True : Me.TextBox2.Enabled = True
            Me.Button3.Enabled = True
        End If
        If Frm_Movimientos_Exp.ComboBox1.Text = "BAJAS" Then
            Me.Text = "Agregar Registro <BAJAS>"
            Me.ComboBox8.Enabled = False : Me.TextBox2.Enabled = False
            Me.Button3.Enabled = False
        End If
        If Frm_Movimientos_Exp.ComboBox1.Text = "TRASLADOS" Then
            Me.Text = "Agregar Registro <TRASLADOS>"
            Me.ComboBox8.Enabled = False : Me.TextBox2.Enabled = False
            Me.Button3.Enabled = True
        End If
        RemoveHandler Me.ComboBox5.SelectedIndexChanged, AddressOf ComboBox5_SelectedIndexChanged
        Me.TextBox5.Text = "0000"
        Me.CheckBox1.Enabled = False
        Me.CheckBox1.Checked = True
        Me.CheckBox2.Checked = True
        Me.TextBox1.Text = ""
        Me.TextBox6.Text = ""
        Me.TextBox7.Text = ""
        Me.TextBox8.Text = ""
        Me.TextBox4.Text = ""
        Call CARGAR_COMBO_CAT_DE_SUB_TIPO_DE_MOVIMIENTO_LEGAL_DE_EXP()
        Call CARGAR_COMBO_CAT_DE_SUB_TIPO_DE_MOVIMIENTO_REAL_DE_EXP()
        Call Clases.BUSCAR_FECHA_Y_HORA()
        Me.DateTimePicker1.Value = FECHA_SERVIDOR
        Me.DateTimePicker2.Value = FECHA_SERVIDOR
        Call CARGAR_COMBO_CAT_DE_CARGOS_DE_ESTRUCTURA()
        Call CARGAR_COMBO_CAT_DE_GRADO_PLANTILLA()
        Call CARGAR_COMBO_CAT_DE_GRADO_REAL()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_1()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_2()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_3()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_4()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_5()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_6()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_7()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_8()
        Me.TextBox5.Text = ""
        Call CARGAR_COMBO_CAT_DE_NOMINA()
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.ComboBox6.Focus()
        AddHandler Me.ComboBox5.SelectedIndexChanged, AddressOf ComboBox5_SelectedIndexChanged
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_NOMINA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE NOMINA] ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox8
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_NOMINA"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_8()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTRUCTURA NIVEL 8] ORDER BY ORDEN", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count <> 0 Then
                With Me.ComboBox14
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_N8"
                End With
            Else
                Me.ComboBox14.Text = ""
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_7()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTRUCTURA NIVEL 7] ORDER BY ORDEN", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count <> 0 Then
                With Me.ComboBox13
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_N7"
                End With
            Else
                Me.ComboBox13.Text = ""
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_6()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTRUCTURA NIVEL 6] ORDER BY ORDEN", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count <> 0 Then
                With Me.ComboBox12
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_N6"
                End With
            Else
                Me.ComboBox12.Text = ""
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_5()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTRUCTURA NIVEL 5] ORDER BY ORDEN", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count <> 0 Then
                With Me.ComboBox11
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_N5"
                End With
            Else
                Me.ComboBox11.Text = ""
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_4()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTRUCTURA NIVEL 4] ORDER BY ORDEN", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count <> 0 Then
                With Me.ComboBox10
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_N4"
                End With
            Else
                Me.ComboBox10.Text = ""
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_3()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTRUCTURA NIVEL 3] ORDER BY ORDEN", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count <> 0 Then
                With Me.ComboBox9
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_N3"
                End With
            Else
                Me.ComboBox9.Text = ""
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_2()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTRUCTURA NIVEL 2] ORDER BY ORDEN", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count <> 0 Then
                With Me.ComboBox7
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_N2"
                End With
            Else
                Me.ComboBox7.Text = ""
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_1()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTRUCTURA NIVEL 1] ORDER BY ORDEN", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox5
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_N1"
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
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE GRADO REAL] ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox4
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
    Private Sub CARGAR_COMBO_CAT_DE_GRADO_PLANTILLA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE GRADO PLANTILLA] ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox3
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
    Private Sub CARGAR_COMBO_CAT_DE_CARGOS_DE_ESTRUCTURA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE CARGOS DE ESTRUCTURA] ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox2
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
    Private Sub CARGAR_COMBO_CAT_DE_SUB_TIPO_DE_MOVIMIENTO_REAL_DE_EXP()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE SUB TIPO DE MOVIMIENTO REAL DE EXP] WHERE ID_T_M_EXP = " & Frm_Movimientos_Exp.ComboBox1.SelectedValue & " ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_ST_M_R_EXP"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_SUB_TIPO_DE_MOVIMIENTO_LEGAL_DE_EXP()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE SUB TIPO DE MOVIMIENTO LEGAL DE EXP] WHERE ID_T_M_EXP = " & Frm_Movimientos_Exp.ComboBox1.SelectedValue & " ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox6
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_ST_M_L_EXP"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_2()
    End Sub
    Private Sub ComboBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox6.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            If Me.ComboBox6.Text = "NUEVO INGRESO" Or Me.ComboBox6.Text = "TRASLADO" Or Me.ComboBox6.Text = "REINGRESO" Then
                Me.ComboBox1.Text = Me.ComboBox6.Text
            End If
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
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
    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            Me.ComboBox4.Text = Me.ComboBox3.Text
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
    Private Sub ComboBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox7.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            If Me.TextBox5.Text = "" Then
                Me.TextBox5.Text = "0000"
            End If
            Me.TextBox5.Text = Format(CInt(Me.TextBox5.Text), "0000")
            If Frm_Movimientos_Exp.ComboBox1.Text = "ALTAS" Or Frm_Movimientos_Exp.ComboBox1.Text = "BAJAS" Or Frm_Movimientos_Exp.ComboBox1.Text = "TRASLADOS" Then
                Call BUSCAR_DATOS_CON_N_ORDEN()
                If DATOS_CON_N_ORDEN = False Then
                    MsgBox("El No. de Orden Mixto no es válido, Cargo no esta vacante (ALTAS), esta Ocupado (BAJA) o no esta Asignado al Usuario Actual", vbInformation, "Mensaje del Sistema")
                    Me.TextBox8.Text = ""
                    Me.TextBox4.Text = ""
                    Me.TextBox2.Text = "0"
                    Me.TextBox5.SelectAll()
                    Me.TextBox5.Focus()
                    Exit Sub
                End If
            End If
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox5_DoubleClick(sender As Object, e As EventArgs) Handles TextBox5.DoubleClick
        Frm_Movimientos_Exp_Add_Buscar_Norden.ShowDialog()
        If CERRAR = False Then
            Me.TextBox5.Text = maOM
            If Me.TextBox5.Text = "" Then
                Me.TextBox5.Text = "0000"
            End If
            Me.TextBox5.Text = Format(CInt(Me.TextBox5.Text), "0000")
            If Frm_Movimientos_Exp.ComboBox1.Text = "ALTAS" Or Frm_Movimientos_Exp.ComboBox1.Text = "BAJAS" Or Frm_Movimientos_Exp.ComboBox1.Text = "TRASLADOS" Then
                Call BUSCAR_DATOS_CON_N_ORDEN()
                If DATOS_CON_N_ORDEN = False Then
                    MsgBox("El No. de Orden Mixto no es válido, Cargo no esta vacante (ALTAS), esta Ocupado (BAJA) o no esta Asignado al Usuario Actual", vbInformation, "Mensaje del Sistema")
                    Me.TextBox8.Text = ""
                    Me.TextBox4.Text = ""
                    Me.TextBox2.Text = "0"
                    Me.TextBox5.SelectAll()
                    Me.TextBox5.Focus()
                    Exit Sub
                End If
            End If
        End If
    End Sub
    Dim DATOS_CON_N_ORDEN As Boolean
    Dim vMIXTA, vMILITAR, vPAME As Boolean
    Private Sub BUSCAR_DATOS_CON_N_ORDEN()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            If Frm_Movimientos_Exp.ComboBox1.Text = "ALTAS" Or Frm_Movimientos_Exp.ComboBox1.Text = "TRASLADOS" Then
                If isuID_TIPO_USUARIO = 5 Then
                    CADENAsql = "SELECT * FROM [dbo].[VISTA MAESTRO DE CARGOS SIT] WHERE N_ORDEN_MIXTA = '" & Me.TextBox5.Text & "' AND ID_SITUACION = 2 AND D_ADMIN = '" & isuUSUARIO & "'"
                Else
                    CADENAsql = "SELECT * FROM [dbo].[VISTA MAESTRO DE CARGOS SIT] WHERE N_ORDEN_MIXTA = '" & Me.TextBox5.Text & "' AND ID_SITUACION = 2"
                End If
                Me.ComboBox8.Enabled = True
                Me.TextBox2.Enabled = True
            End If
            If Frm_Movimientos_Exp.ComboBox1.Text = "BAJAS" Then
                Me.ComboBox8.Enabled = False
                Me.TextBox2.Enabled = False
                If isuID_TIPO_USUARIO = 5 Then
                    CADENAsql = "SELECT * FROM [dbo].[VISTA MAESTRO DE CARGOS SIT] WHERE N_ORDEN_MIXTA = '" & Me.TextBox5.Text & "' AND ID_SITUACION = 1 AND D_ADMIN = '" & isuUSUARIO & "'"
                Else
                    CADENAsql = "SELECT * FROM [dbo].[VISTA MAESTRO DE CARGOS SIT] WHERE N_ORDEN_MIXTA = '" & Me.TextBox5.Text & "' AND ID_SITUACION = 1"
                End If
            End If
            MiDataAdapter = New SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                Me.TextBox8.Text = MiDataTable.Rows(0).Item(40).ToString
                Me.TextBox4.Text = MiDataTable.Rows(0).Item(41).ToString
                Me.ComboBox2.Text = MiDataTable.Rows(0).Item(32).ToString
                Me.ComboBox3.Text = MiDataTable.Rows(0).Item(34).ToString
                Me.ComboBox4.Text = MiDataTable.Rows(0).Item(36).ToString
                Me.ComboBox5.Text = MiDataTable.Rows(0).Item(4).ToString
                Me.ComboBox7.Text = MiDataTable.Rows(0).Item(7).ToString
                Me.ComboBox9.Text = MiDataTable.Rows(0).Item(10).ToString
                Me.ComboBox10.Text = MiDataTable.Rows(0).Item(13).ToString
                Me.ComboBox11.Text = MiDataTable.Rows(0).Item(16).ToString

                Me.ComboBox12.Text = MiDataTable.Rows(0).Item(19).ToString
                Me.ComboBox13.Text = MiDataTable.Rows(0).Item(22).ToString
                Me.ComboBox14.Text = MiDataTable.Rows(0).Item(25).ToString

                Me.TextBox7.Text = MiDataTable.Rows(0).Item(0).ToString
                Me.TextBox6.Text = MiDataTable.Rows(0).Item(39).ToString
                vMIXTA = MiDataTable.Rows(0).Item(46).ToString
                vMILITAR = MiDataTable.Rows(0).Item(47).ToString
                vPAME = MiDataTable.Rows(0).Item(48).ToString
                Call BUSCAR_NOMINA
                DATOS_CON_N_ORDEN = True

            Else
                Me.ComboBox2.Text = ""
                Me.ComboBox3.Text = ""
                Me.ComboBox4.Text = ""
                Me.ComboBox5.Text = ""
                Me.ComboBox7.Text = ""
                Me.ComboBox9.Text = ""
                Me.ComboBox10.Text = ""
                Me.ComboBox11.Text = ""
                Me.ComboBox12.Text = ""
                Me.ComboBox13.Text = ""
                Me.ComboBox14.Text = ""
                Me.TextBox2.Text = "0"
                Me.TextBox7.Text = "0"
                Me.TextBox6.Text = "0"
                DATOS_CON_N_ORDEN = False
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub BUSCAR_NOMINA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            CADENAsql = "SELECT * FROM [dbo].[VISTA MAESTRO DE NOMINAS Y SALARIOS] WHERE ID_M_P = " & Me.TextBox6.Text & " AND ACTIVO = 'TRUE'"
            MiDataAdapter = New SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                Me.ComboBox8.Text = MiDataTable.Rows(0).Item(6).ToString
                Me.TextBox2.Text = MiDataTable.Rows(0).Item(8).ToString
            Else
                Me.ComboBox8.Text = "PAME"
                Me.TextBox2.Text = "0"
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ComboBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox8.KeyDown
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
    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Frm_Movimientos_Exp.ComboBox1.Text = "TRASLADOS" Then
            Frm_Movimientos_Exp_Add_Seleccionar_Recurso.ShowDialog()
            If CERRAR = False Then
                Me.TextBox8.Text = cEMPLEADO
                Me.TextBox4.Text = bAN
                Me.TextBox6.Text = vID_M_P
            End If
        End If
        If Frm_Movimientos_Exp.ComboBox1.Text = "ALTAS" Then
            If Me.ComboBox6.Text = "NUEVO INGRESO" Then
                Frm_Movimientos_Exp_Add_Recurso_Nuevo.ShowDialog()
                If CERRAR = False Then
                    Me.TextBox8.Text = cEMPLEADO
                    Me.TextBox4.Text = bAN
                    Me.TextBox6.Text = vID_M_P
                    Me.DateTimePicker2.Value = cFECHAING
                End If
            End If
            If Me.ComboBox6.Text = "REINGRESO" Then
                Frm_Movimientos_Exp_Add_Seleccionar_Recurso_Reingreso.ShowDialog()
                If CERRAR = False Then
                    Me.TextBox8.Text = cEMPLEADO
                    Me.TextBox4.Text = bAN
                    Me.TextBox6.Text = vID_M_P
                    Me.DateTimePicker2.Value = cFECHAING
                End If
            End If
        End If
    End Sub
    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        If Me.ComboBox6.Text = "NUEVO INGRESO" Or Me.ComboBox6.Text = "TRASLADO" Or Me.ComboBox6.Text = "REINGRESO" Then
            Me.ComboBox1.Text = Me.ComboBox6.Text
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'If Len(Me.TextBox1.Text) = 0 Or Me.TextBox1.Text = "0" Or Me.TextBox1.Text = "" Then
        'MsgBox("No se encuentra el ID del Registro", vbInformation, "Mensaje del Sistema")
        'Me.Button3.Focus()
        'Exit Sub
        'End If

        If Len(Me.TextBox6.Text) = 0 Or Me.TextBox6.Text = "0" Or Me.TextBox6.Text = "1" Then
            MsgBox("Debe seleccionar a la Persona", vbInformation, "Mensaje del Sistema")
            Me.Button3.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox6.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar la Causa Legal del Movimiento Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox6.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar la Causa Real del Movimiento Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        Call Clases.BUSCAR_FECHA_Y_HORA()
        If Me.DateTimePicker1.Value.Year < FECHA_SERVIDOR.Year Then
            Dim MENSAJE_1 = MsgBox("El Año del Movimiento es menor al año actual, ¿Esta seguro de Continuar?", vbInformation + vbYesNo, "Mensaje del Sistema")
            If MENSAJE_1 = vbNo Then
                Me.DateTimePicker1.Focus()
                Exit Sub
            End If
        End If
        If Len(Me.ComboBox2.SelectedValue) = 0 Then
            MsgBox("No se encuentra el Cargo para el No. de Orden actual", vbInformation, "Mensaje del Sistema")
            Me.TextBox5.SelectAll()
            Me.TextBox5.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox3.SelectedValue) = 0 Then
            MsgBox("No se encuentra el Grado por Plantilla para el No. de Orden actual", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        If Len(Me.ComboBox4.SelectedValue) = 0 Then
            MsgBox("No se encuentra el Grado Real para el No. de Orden actual", vbInformation, "Mensaje del Sistema")
            Me.TextBox5.SelectAll()
            Me.TextBox5.Focus()
            Exit Sub
        End If
        Me.TextBox5.Text = Format(CInt(Me.TextBox5.Text), "0000")
        If Len(Me.TextBox5.Text) <> 4 Or Me.TextBox5.Text = "0000" Or Me.TextBox5.Text = "" Then
            MsgBox("Debe Digitar el No. de Orden Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox5.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox8.Text) = 0 Then
            MsgBox("No se encuentra el codigo del Empleado", vbInformation, "Mensaje del Sistema")
            Me.Button3.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox4.Text) = 0 Then
            MsgBox("No se encuentran los Apellidos y Nombres del Empleado", vbInformation, "Mensaje del Sistema")
            Me.Button3.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox8.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar la Nómina correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox8.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox2.Text) = 0 Or Me.TextBox2.Text = "" Or Me.TextBox2.Text = "0" Then
            MsgBox("Debe digitar el monto del Salario del Empleado", vbInformation, "Mensaje del Sistema")
            Me.TextBox2.Focus()
            Exit Sub
        End If
        Call GENERAR_CODIGO()
        If TextBox1.Text = "0" Or TextBox1.Text = "" Then
            MsgBox("No se ha generado el Codigo del registro", vbInformation, "Mensaje del Sistema")
            Me.TextBox1.Text = "0"
            Me.Button2.Focus()
            Exit Sub
        End If
        If Frm_Movimientos_Exp.ComboBox1.Text = "ALTAS" Then                    'VALIDAR FECHAS EN MOV. ALTA
            If Mid(Me.DateTimePicker1.Value, 1, 10) <> Mid(Me.DateTimePicker2.Value, 1, 10) Then
                Dim MENSAJE_2 = MsgBox("El Movimiento de <ALTA> no deberia tener una Fecha de Grabacion diferente a la Fecha de Aplicacion, ¿Esta seguro de Continuar?", vbInformation + vbYesNo, "Mensaje del Sistema")
                If MENSAJE_2 = vbNo Then
                    Me.DateTimePicker1.Focus()
                    Exit Sub
                End If
            End If
        End If
        If Frm_Movimientos_Exp.ComboBox1.Text = "BAJAS" Then                    'VALIDAR FECHAS EN MOV. BAJAS, FECHA DE GRABACION DEBE SER
            If Me.DateTimePicker1.Value > Me.DateTimePicker2.Value Then         'MENOR LA FECHA DE APLICACION
                Dim MENSAJE_2 = MsgBox("La Fecha de Aplicacion no deberia ser menor a la Fecha de Grabacion, ¿Esta seguro de Continuar?", vbInformation + vbYesNo, "Mensaje del Sistema")
                If MENSAJE_2 = vbNo Then
                    Me.DateTimePicker1.Focus()
                    Exit Sub
                End If
            End If
        End If
        If Frm_Movimientos_Exp.ComboBox1.Text = "BAJAS" Then                    'VALIDAR CHECK DE APLICACION
            If Me.DateTimePicker1.Value <> Me.DateTimePicker2.Value Then
                If Me.DateTimePicker1.Value < Me.DateTimePicker2.Value Then
                    Me.CheckBox1.Checked = False
                    Me.CheckBox2.Checked = False
                End If
            End If
        End If
        If Frm_Movimientos_Exp.ComboBox1.Text = "TRASLADOS" Then               'VALIDAR FECHAS EN MOV. TRASLADO
            If Me.DateTimePicker1.Value <> Me.DateTimePicker2.Value Then
                Dim MENSAJE_2 = MsgBox("El Movimiento de <TRASLADO> no deberia tener una Fecha de Grabacion diferente a la Fecha de Aplicacion, ¿Esta seguro de Continuar?", vbInformation + vbYesNo, "Mensaje del Sistema")
                If MENSAJE_2 = vbNo Then
                    Me.DateTimePicker1.Focus()
                    Exit Sub
                End If
            End If
        End If

        Call GRABAR_REGISTRO_DE_HISTORICO_PERSONA()
        '-------------------
        If Frm_Movimientos_Exp.ComboBox1.Text = "ALTAS" Then
            vID_ESTADO_P = 1
            Call ACTUALIZAR_ID_ESTADO_P_EN_MP()
            Call ACTUALIZAR_REGISTRO_EN_MC()
            Call DESACTIVAR_NOMINA_SELECCIONADA()
            Call GENERAR_CODIGO_MAESTRO_DE_NOMINAS_Y_SALARIOS()
            Call INGRESAR_NOMINA_Y_SALARIO()
        End If
        If Frm_Movimientos_Exp.ComboBox1.Text = "BAJAS" Then
            If Me.DateTimePicker1.Value = Me.DateTimePicker2.Value Then    'SI LAS FECHAS SON IGUALES SE DEBE APLICAR BAJA
                vID_ESTADO_P = 2
                Call ACTUALIZAR_ID_ESTADO_P_EN_MP()
                Call ACTUALIZAR_REGISTRO_EN_MC()
            End If
        End If
        If Frm_Movimientos_Exp.ComboBox1.Text = "TRASLADOS" Then
            Call ACTUALIZAR_REGISTRO_EN_MC()
        End If
        vID_HIST_M = Me.TextBox1.Text
        CERRAR = False
        Me.Close()
    End Sub
    Private Sub INGRESAR_NOMINA_Y_SALARIO()
        Dim fFEC_MOV As String = Me.DateTimePicker1.Text
        Dim fechaMovimiento = "Convert(VARCHAR(10), Convert(Date, '" + fFEC_MOV + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE NOMINAS Y SALARIOS] (ID_M_NYS, ID_M_P, ID_MOVIMIENTO_N, F_MOVIMIENTO_N, ID_NOMINA, PROPUESTA, SALARIO, SAL_MINIMO, OBSERVACIONES, SAL_ACTUAL)" &
                                  "values (" & CInt(vID_M_NYS) & ", " & Me.TextBox6.Text & ", 2, " & fechaMovimiento & ", " & Me.ComboBox8.SelectedValue & ", '" & Me.TextBox2.Text & "', '" & Me.TextBox2.Text & "', 'FALSE', 'INGRESO AUTOMATICO POR ALTA', 'TRUE')"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim vID_M_NYS As Integer
    Private Sub GENERAR_CODIGO_MAESTRO_DE_NOMINAS_Y_SALARIOS()
        CODIGO = 0
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            MiDataAdapter = New SqlDataAdapter("SELECT ID_M_NYS FROM [MAESTRO DE NOMINAS Y SALARIOS] ORDER BY ID_M_NYS", Conexion.CxRRHH)
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
            vID_M_NYS = CODIGO
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub DESACTIVAR_NOMINA_SELECCIONADA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE NOMINAS Y SALARIOS] SET SAL_ACTUAL = 'FALSE' WHERE ID_M_P = " & CInt(Me.TextBox6.Text) & " AND ID_NOMINA = " & Me.ComboBox8.SelectedValue & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ACTUALIZAR_REGISTRO_EN_MC()
        'AGREGAR MOVIMIENTO DE ALTAS
        If Frm_Movimientos_Exp.ComboBox1.Text = "ALTAS" Then
            CADENAsql5 = "UPDATE [dbo].[MAESTRO DE CARGOS] SET ID_M_P = " & Me.TextBox6.Text & ", ID_SITUACION = 1, ID_GR = " & Me.ComboBox4.SelectedValue & " WHERE ID_M_C = " & Me.TextBox7.Text & ""
            Call ACTUALIZAR_MC()
            'AGREGAR CONTRATOS
            Call BUSCAR_SI_EXISTE_REGISTRO_EN_CONTRATOS
            If EXISTE_REGISTRO_EN_CONTRATOS = True Then
                Call BORRAR_CONTRATOS
            End If
            Call AGREGAR_REGISTRO_MCyE_ALTA()
        End If
        'AGREGAR MOVIMIENTO DE BAJAS
        If Frm_Movimientos_Exp.ComboBox1.Text = "BAJAS" Then
            CADENAsql5 = "UPDATE [dbo].[MAESTRO DE CARGOS] SET ID_M_P = 1, ID_SITUACION = 2, ID_GR = " & Me.ComboBox4.SelectedValue & " WHERE ID_M_C = " & Me.TextBox7.Text & ""
            Call ACTUALIZAR_MC()
        End If
        'AGREGAR MOVIMIENTO DE TRASLADOS
        If Frm_Movimientos_Exp.ComboBox1.Text = "TRASLADOS" Then
            'ACTUALIZAR EL CARGO ORIGEN A VACANTE
            CADENAsql5 = "UPDATE [dbo].[MAESTRO DE CARGOS] SET ID_GR = " & Me.ComboBox4.SelectedValue & ", ID_M_P = 1, ID_SITUACION = 2 WHERE ID_M_C = " & cID_M_C & ""
            Call ACTUALIZAR_MC()
            'TRASLADAR
            CADENAsql5 = "UPDATE [dbo].[MAESTRO DE CARGOS] SET ID_GR = " & Me.ComboBox4.SelectedValue & ", ID_M_P = " & Me.TextBox6.Text & ", ID_SITUACION = 1 WHERE ID_M_C = " & Me.TextBox7.Text & ""
            Call ACTUALIZAR_MC()
        End If
    End Sub
    Dim ceFI As String
    Dim ceFF As String
    Dim tipo As Integer
    Dim firma As Boolean

    Dim ceI1CONT, ceF1CONT, ceI1EVAL, ceF1EVAL, ceI2CONT, ceF2CONT, ceI2EVAL, ceF2EVAL, ceF3CONT As Object
    Private Sub AGREGAR_REGISTRO_MCyE_ALTA()
        '1 CONTRATO
        ceI1CONT = Mid(Me.DateTimePicker1.Value, 1, 10)     'IGUAL A LA FECHA DE INGRESO
        ceF1CONT = Mid(Me.DateTimePicker1.Value.AddMonths(3), 1, 10)
        ceF1CONT = CDate(ceF1CONT).AddDays(-1)              '3 MESES + 1 DIA DESPUES DE LA FECHA DE INGRESO
        tipo = 1
        ceFI = ceI1CONT
        ceFF = ceF1CONT
        firma = True
        Call GENERAR_CODIGO_CYE()
        Call AGREGAR_REGISTRO_CYE()
        '1 EVALUACION
        ceI1EVAL = Mid(Me.DateTimePicker1.Value.AddMonths(2), 1, 10)    '2 MESES DESPUES DE LA FECHA DE INGRESO
        ceF1EVAL = CDate(ceI1EVAL).AddMonths(1)                         '1 MES DESPUES DE INICIAR LA 1 EVALUACION
        tipo = 2
        ceFI = ceI1EVAL
        ceFF = ceF1EVAL
        firma = False
        Call GENERAR_CODIGO_CYE()
        Call AGREGAR_REGISTRO_CYE()
        '2 CONTRATO
        ceI2CONT = CDate(ceF1CONT).AddDays(1)               '1 DIA DESPUES DE FINALIZAR EL PRIMER CONTRATO
        ceF2CONT = CDate(ceF1CONT).AddMonths(6)             '6 MESES DESPUES DE FINALIZAR EL PRIMER CONTRATO
        tipo = 3
        ceFI = ceI2CONT
        ceFF = ceF2CONT
        firma = False
        Call GENERAR_CODIGO_CYE()
        Call AGREGAR_REGISTRO_CYE()
        '2 EVALUACION
        ceI2EVAL = CDate(ceI2CONT).AddMonths(4)             '4 MESES DESPUES DE INICIAR LA PRIMERA EVALUACION
        ceF2EVAL = CDate(ceI2EVAL).AddMonths(1)             '1 MES DESPUES DE HABER INICIADO LA SEGUNDA EVALUACION
        tipo = 4
        ceFI = ceI2EVAL
        ceFF = ceF2EVAL
        firma = False
        Call GENERAR_CODIGO_CYE()
        Call AGREGAR_REGISTRO_CYE()
        'CONTRATO INDEFINIDO
        ceF3CONT = CDate(ceF2CONT).AddDays(1)               '1 DIA  DESPUES DE HABER FINALIZADO EL SEGUNDO CONTRATO
        tipo = 5
        ceFI = ceF3CONT
        ceFF = ""
        firma = False
        Call GENERAR_CODIGO_CYE()
        Call AGREGAR_REGISTRO_CYE()
    End Sub
    Private Sub AGREGAR_REGISTRO_CYE()
        Dim fFECHA_I As String
        If ceFI <> "" Then
            fFECHA_I = Mid(ceFI, 1, 10)
        Else
            fFECHA_I = "NULL"
        End If
        Dim f1 As Object = If(fFECHA_I <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_I + "', 105), 23)", "NULL")

        Dim fFECHA_F As String
        If ceFF <> "" Then
            fFECHA_F = Mid(ceFF, 1, 10)
        Else
            fFECHA_F = "NULL"
        End If
        Dim f2 As Object = If(fFECHA_F <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_F + "', 105), 23)", "NULL")

        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE CONTRATOS Y EVALUACIONES] (ID_M_CYE, ID_M_P, ID_CONT, FECHA_I, FECHA_F, FIRMADO, ID_C_E, OBSERVACIONES)" &
                                  "values (" & CInt(vID_M_CYE) & ", " & Me.TextBox6.Text & ", " & tipo & ", " & f1 & ", " & f2 & ", '" & firma & "', 1, 'INGRESO AUTOMATICO DESDE ALTA')"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim vID_M_CYE As Integer
    Private Sub GENERAR_CODIGO_CYE()
        CODIGO = 0
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            MiDataAdapter = New SqlDataAdapter("SELECT ID_M_CYE FROM [MAESTRO DE CONTRATOS Y EVALUACIONES] ORDER BY ID_M_CYE", Conexion.CxRRHH)
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
            vID_M_CYE = CODIGO
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub BORRAR_CONTRATOS()
        Me.Cursor = Cursors.WaitCursor
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE CONTRATOS Y EVALUACIONES] WHERE [ID_M_P] = " & Me.TextBox6.Text & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
        Me.Cursor = Cursors.Default
    End Sub
    Dim EXISTE_REGISTRO_EN_CONTRATOS As Boolean
    Private Sub BUSCAR_SI_EXISTE_REGISTRO_EN_CONTRATOS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [MAESTRO DE CONTRATOS Y EVALUACIONES] WHERE ID_M_P = " & Me.TextBox6.Text & "", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                EXISTE_REGISTRO_EN_CONTRATOS = True
            Else
                EXISTE_REGISTRO_EN_CONTRATOS = False
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ACTUALIZAR_MC()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = CADENAsql5
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim vID_ESTADO_P As Integer
    Private Sub ACTUALIZAR_ID_ESTADO_P_EN_MP()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE PERSONAS] SET ID_ESTADO_P = " & vID_ESTADO_P & " WHERE ID_M_P = " & CInt(Me.TextBox6.Text) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub GRABAR_REGISTRO_DE_HISTORICO_PERSONA()
        Dim OBS As String = "REPLACE(REPLACE(REPLACE(convert(nvarchar(max),'" & Me.TextBox3.Text & "'),Char(10),' '), Char(13),' '), ' ',' ')"
        cadena_CAMPOS1 = "ID_HIST_M, ID_M_P, FEC_MOV, FEC_APLICA, ID_T_M_EXP, ID_ST_M_L_EXP, " &
                 "ID_ST_M_R_EXP, D_CARGO_ES, D_GP, D_GR, D_N1, " &
                 "D_N2, D_N3, D_N4, D_N5, D_N6, D_N7, D_N8, N_ORDEN, " &
                 "D_NOMINA, SALARIO, DETALLE, MIXTA, MILITAR, " &
                 "PAME, APLICAR, ESTADISTICO, USUARIO_ACT, FECHA_ACT"

        Dim fFEC_MOV As String = Mid(Me.DateTimePicker1.Value, 1, 10)
        Dim fechaMovimiento = "Convert(VARCHAR(10), Convert(Date, '" + fFEC_MOV + "', 105), 23)"
        Dim fFEC_APLI As String = Mid(Me.DateTimePicker2.Value, 1, 10)
        Dim fechaaplicacion = "Convert(VARCHAR(10), Convert(Date, '" + fFEC_APLI + "', 105), 23)"

        Call Clases.BUSCAR_FECHA_Y_HORA()
        Dim fFEC_AUDI As String = Mid(FECHA_SERVIDOR, 1, 10)
        Dim fechaAudi = "Convert(VARCHAR(10), Convert(Date, '" + fFEC_AUDI + "', 105), 23)"

        cadena_VALORES1 = "" & Me.TextBox1.Text & ", " & Me.TextBox6.Text & ", " & fechaMovimiento & ", " & fechaaplicacion & ", " & Frm_Movimientos_Exp.ComboBox1.SelectedValue & ", " & Me.ComboBox6.SelectedValue & ", " &
                 "" & Me.ComboBox1.SelectedValue & ", '" & Me.ComboBox2.Text & "', '" & Me.ComboBox3.Text & "', '" & Me.ComboBox4.Text & "', '" & Me.ComboBox5.Text & "', " &
                 "'" & Me.ComboBox7.Text & "', '" & Me.ComboBox9.Text & "', '" & Me.ComboBox10.Text & "', '" & Me.ComboBox11.Text & "', '" & Me.ComboBox12.Text & "', '" & Me.ComboBox13.Text & "', '" & Me.ComboBox14.Text & "', '" & Me.TextBox5.Text & "', " &
                 "'" & Me.ComboBox8.Text & "', '" & Me.TextBox2.Text & "', " & OBS & ", '" & vMIXTA & "', '" & vMILITAR & "', " &
                 "'" & vPAME & "', '" & Me.CheckBox1.Checked & "', '" & Me.CheckBox2.Checked & "',  '" & isuCUENTA & "', " & fechaAudi & ""

        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim SQL As String
        Try
            SQL = "INSERT INTO [HISTORICO DE MOVIMIENTO] (" & cadena_CAMPOS1 & ") values (" & cadena_VALORES1 & ")"
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
    Dim cadena_CAMPOS1, cadena_VALORES1 As String
    Private Sub GRABAR_REGISTRO_DE_PERSONA()


    End Sub
    Private Sub GENERAR_CODIGO()
        Dim CODIGO As Integer
        CODIGO = 0
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            MiDataAdapter = New SqlDataAdapter("SELECT ID_HIST_M FROM [HISTORICO DE MOVIMIENTO] ORDER BY ID_HIST_M", Conexion.CxRRHH)
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
End Class