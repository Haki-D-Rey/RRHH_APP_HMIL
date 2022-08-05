Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Public Class Frm_Expedientes_Activos
    Private Sub Button7_Click(sender As Object, e As System.EventArgs) Handles Button7.Click
        EXPEDIENTES = False
        Me.TextBox35.Text = ""
        Me.TextBox36.Text = ""
        vID_M_P = 0 : vID_M_C = 0
        Frm_Principal.Timer1.Start()
        Me.Close()
    End Sub
    Private Sub Frm_Expedientes_Activos_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            EXPEDIENTES = False
            Me.TextBox35.Text = ""
            Me.TextBox36.Text = ""
            vID_M_P = 0 : vID_M_C = 0
            Frm_Principal.Timer1.Start()
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Expedientes_Activos_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Frm_Principal.Timer1.Stop()
        EXPEDIENTES = True
        Dim tp As TabPage = TabGeneral.TabPages(1)
        TabGeneral.SelectedTab = tp
        Me.Label4.Text = Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Text
        vID_M_P = 0 : vID_M_C = 0
        Call NO_ACTIVAR()
        Call SELECCIONAR_CADENA()
        Me.TextBox35.Text = ""
        Me.TextBox36.Text = ""
        TabGeneral.Focus()
        Me.Show()
        '-------------------------------------------------
        Frm_Expedientes_Buscar_Activos.ShowDialog()
        If CERRAR = False Then
            tp = TabGeneral.TabPages(1)
            TabGeneral.SelectedTab = tp
            Me.Label4.Text = Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Text
            Call NO_ACTIVAR()
            Call SELECCIONAR_CADENA()
            Call DESACTIVAR_TODO()
            TabGeneral.Focus()
            '------------------------------------------------------------------------------------
            Call BUSCAR_UBICACION()
            '------------------------------------------------------------------------------------
        End If
        '-------------------------------------------------
    End Sub
    Private Sub BUSCAR_UBICACION()
        Dim n1, n2, n3, n4, n5, n6, n7, n8, cargo, gradoreal As String
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [VISTA MAESTRO DE CARGOS] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_C", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                Me.Cursor = Cursors.WaitCursor
                If Not IsDBNull(MiDataTable.Rows(0).Item(4).ToString) Then : n1 = MiDataTable.Rows(0).Item(4).ToString : Else : n1 = "" : End If
                If Not IsDBNull(MiDataTable.Rows(0).Item(7).ToString) Then : n2 = MiDataTable.Rows(0).Item(7).ToString : Else : n2 = "" : End If
                If Not IsDBNull(MiDataTable.Rows(0).Item(10).ToString) Then : n3 = MiDataTable.Rows(0).Item(10).ToString : Else : n3 = "" : End If
                If Not IsDBNull(MiDataTable.Rows(0).Item(13).ToString) Then : n4 = MiDataTable.Rows(0).Item(13).ToString : Else : n4 = "" : End If
                If Not IsDBNull(MiDataTable.Rows(0).Item(16).ToString) Then : n5 = MiDataTable.Rows(0).Item(16).ToString : Else : n5 = "" : End If
                If Not IsDBNull(MiDataTable.Rows(0).Item(19).ToString) Then : n6 = MiDataTable.Rows(0).Item(19).ToString : Else : n6 = "" : End If
                If Not IsDBNull(MiDataTable.Rows(0).Item(22).ToString) Then : n7 = MiDataTable.Rows(0).Item(22).ToString : Else : n7 = "" : End If
                If Not IsDBNull(MiDataTable.Rows(0).Item(25).ToString) Then : n8 = MiDataTable.Rows(0).Item(25).ToString : Else : n8 = "" : End If

                If Not IsDBNull(MiDataTable.Rows(0).Item(32).ToString) Then : cargo = MiDataTable.Rows(0).Item(32).ToString : Else : cargo = "" : End If
                If Not IsDBNull(MiDataTable.Rows(0).Item(36).ToString) Then : gradoreal = MiDataTable.Rows(0).Item(36).ToString : Else : gradoreal = "" : End If
                '------------------------------------------------------------------------------------
                Dim cadenaUBIC As String
                cadenaUBIC = n1
                If n2 = "" Then
                    GoTo SALIR
                Else
                    cadenaUBIC = cadenaUBIC & "\ " & n2
                End If
                If n3 = "" Then
                    GoTo SALIR
                Else
                    cadenaUBIC = cadenaUBIC & "\ " & n3
                End If
                If n4 = "" Then
                    GoTo SALIR
                Else
                    cadenaUBIC = cadenaUBIC & "\ " & n4
                End If
                If n5 = "" Then
                    GoTo SALIR
                Else
                    cadenaUBIC = cadenaUBIC & "\ " & n5
                End If
                If n6 = "" Then
                    GoTo SALIR
                Else
                    cadenaUBIC = cadenaUBIC & "\ " & n6
                End If
                If n7 = "" Then
                    GoTo SALIR
                Else
                    cadenaUBIC = cadenaUBIC & "\ " & n7
                End If
                If n8 = "" Then
                    GoTo SALIR
                Else
                    cadenaUBIC = cadenaUBIC & "\ " & n8
                End If
SALIR:
                Me.TextBox39.Text = cadenaUBIC                                      'UBICACION
                Me.TextBox40.Text = cargo
                Me.TextBox41.Text = gradoreal
                Me.Cursor = Cursors.Default
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub NO_ACTIVAR()
        '1
        Me.TextBox4.ReadOnly = True
        Me.TextBox37.ReadOnly = True
        Me.TextBox38.ReadOnly = True
        Me.ComboBox1.Enabled = False
        Me.ComboBox2.Enabled = False
        Me.ComboBox3.Enabled = False
        Me.ComboBox4.Enabled = False
        Me.CheckBox1.Enabled = False
        Me.ComboBox5.Enabled = False
        Me.ComboBox6.Enabled = False
        Me.ComboBox7.Enabled = False
        Me.ComboBox8.Enabled = False
        Me.ComboBox32.Enabled = False
        Me.CheckBox2.Enabled = False
        Me.CheckBox3.Enabled = False
        Me.CheckBox4.Enabled = False
        Me.Button49.Enabled = True
        Me.Button50.Enabled = False
        Me.Button48.Enabled = False
        '2
        Me.CheckBox5.Enabled = False
        Me.CheckBox6.Enabled = False
        Me.CheckBox7.Enabled = False
        Me.Button2.Enabled = False
        Me.DateTimePicker1.Enabled = False
        Me.DateTimePicker2.Enabled = False
        Me.DateTimePicker3.Enabled = False
        Me.Button2.Enabled = False  'EXAMINAR
        Me.Button4.Enabled = False  'GUARDAR
        Me.Button5.Enabled = False  'EDITAR
        Me.Button8.Enabled = False  'EDITAR
        '3
        Me.TextBox26.ReadOnly = True
        Me.TextBox29.ReadOnly = True
        Me.TextBox43.ReadOnly = True
        Me.DateTimePicker8.Enabled = False
        Me.DateTimePicker9.Enabled = False
    End Sub
    Private Sub TabGeneral_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TabGeneral.Selecting
        Call SELECCIONAR_CADENA()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTABLECIMIENTO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTABLECIMIENTO] order by ID_ESTABLECIMIENTO", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox32
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
    Private Sub CARGAR_COMBO_CAT_DE_ADMINISTRADORES()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ADMINISTRADORES] WHERE ACTIVO = 'TRUE' order by ID_ADMIN", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox33
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_ADMIN"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_IPSS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE IPSS] order by ID_IPSS", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox26
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_IPSS"
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
                With Me.ComboBox36
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_N8"
                End With
            Else
                Me.ComboBox36.Text = ""
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
                With Me.ComboBox35
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_N7"
                End With
            Else
                Me.ComboBox35.Text = ""
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
                With Me.ComboBox34
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_N6"
                End With
            Else
                Me.ComboBox34.Text = ""
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
                With Me.ComboBox31
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_N5"
                End With
            Else
                Me.ComboBox31.Text = ""
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
                With Me.ComboBox30
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_N4"
                End With
            Else
                Me.ComboBox30.Text = ""
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
                With Me.ComboBox29
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_N3"
                End With
            Else
                Me.ComboBox29.Text = ""
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
                With Me.ComboBox2
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_N2"
                End With
            Else
                Me.ComboBox2.Text = ""
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
            With Me.ComboBox1
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
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_ESTRUCTURA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE ESTRUCTURA] order by ID_T_ES", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox3
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
            With Me.ComboBox4
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
            With Me.ComboBox5
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
            With Me.ComboBox6
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
            With Me.ComboBox7
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
            With Me.ComboBox8
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
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_SEXO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE SEXO] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox9
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_T_SEXO"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_TEZ()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TEZ] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox10
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_TEZ"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_CABELLO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE CABELLO] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox11
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_CABELLO"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_NIVEL_ACADEMICO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE NIVEL ACADEMICO] order by ID_N_ACADEMICO", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox13
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_N_ACADEMICO"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_NIVEL_PROFESIONAL()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE NIVEL PROFESIONAL] order by ID_N_PROFESIONAL", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox12
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_N_PROFESIONAL"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_NACIONALIDAD_N()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE NACIONALIDAD_N] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox15
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_NACIONALIDAD_N"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_DPTO_N()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE DPTO_N] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox14
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_DPTO_N"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_MUNICIPIO_N()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE MUNICIPIO_N] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox16
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_M_N"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_NACIONALIDAD_D()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE NACIONALIDAD_D] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox19
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_NACIONALIDAD_D"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_DPTO_D()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE DPTO_D] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox18
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_DPTO_D"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_MUNICIPIO_D()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE MUNICIPIO_D] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox17
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_M_D"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTADO_CIVIL()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTADO CIVIL] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox20
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_E_CIVIL"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_HORARIO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE HORARIO] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox21
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_T_HORARIO"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_AFILIACION()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE AFILIACION] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox23
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_T_AFILIACION"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_PLAZA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE PLAZA] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox22
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_T_PLAZA"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_CATEGORIA_DE_LIC()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE CATEGORIA DE LIC] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox24
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_C_LIC"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_NIVEL_DE_COMPETENCIA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE NIVEL DE COMPETENCIA] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox28
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_N_COMPETENCIA"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_FUNCIONALIDAD()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE FUNCIONALIDAD] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox27
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_F"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_SANGUINEO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO SANGUINEO] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox25
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_T_SANGUINEO"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub LIMPIAR_TEXTOS_1()
        Me.TextBox4.Text = "" : Me.TextBox37.Text = "" : Me.TextBox38.Text = "" : Me.TextBox39.Text = "" : Me.TextBox40.Text = "" : Me.TextBox41.Text = ""
        Me.ComboBox1.Enabled = False : Me.ComboBox2.Enabled = False
        Me.ComboBox3.Enabled = False : Me.ComboBox4.Enabled = False
        Me.CheckBox1.Checked = False
        Me.ComboBox5.Enabled = False : Me.ComboBox6.Enabled = False
        Me.ComboBox7.Enabled = False : Me.ComboBox8.Enabled = False
        Me.CheckBox2.Checked = False
        Me.CheckBox3.Checked = False : Me.CheckBox4.Checked = False
    End Sub
    Private Sub LIMPIAR_TEXTOS_2()
        Me.TextBox5.Text = "" : Me.TextBox6.Text = "" : Me.TextBox30.Text = "" : Me.TextBox9.Text = ""
        Me.TextBox10.Text = "" : Me.TextBox11.Text = "" : Me.TextBox12.Text = "" : Me.TextBox13.Text = ""
        Me.TextBox7.Text = "" : Me.TextBox8.Text = "" : Me.TextBox14.Text = "" : Me.TextBox15.Text = ""
        Me.TextBox16.Text = "" : Me.TextBox17.Text = "" : Me.TextBox18.Text = "" : Me.TextBox21.Text = ""
        Me.TextBox20.Text = "" : Me.TextBox19.Text = "" : Me.TextBox22.Text = "" : Me.TextBox23.Text = ""
        Me.TextBox24.Text = "" : Me.TextBox25.Text = "" : Me.TextBox26.Text = "" : Me.TextBox27.Text = ""
        Me.TextBox31.Text = "" : Me.TextBox32.Text = ""
        Me.TextBox33.Text = "" : Me.TextBox1.Text = "" : Me.TextBox2.Text = "" : Me.TextBox3.Text = ""
        Me.TextBox39.Text = "" : Me.TextBox40.Text = "" : Me.TextBox41.Text = "" : Me.TextBox42.Text = ""
        Me.TextBox34.Text = "" : Me.PictureBox4.Image = Nothing
        Me.TextBox44.Text = "" : Me.TextBox28.Text = ""
    End Sub
    Private Sub SELECCIONAR_CADENA()
        Me.Label4.Text = Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Text
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 1 Then     'INFORMACION ESTRUCTURA
            Call LIMPIAR_TEXTOS_1()
            Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_1()
            Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_2()
            Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_3()
            Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_4()
            Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_5()
            Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_6()
            Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_7()
            Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_8()
            Call CARGAR_COMBO_CAT_DE_TIPO_DE_ESTRUCTURA()
            Call CARGAR_COMBO_CAT_DE_CARGOS_DE_ESTRUCTURA()
            Call CARGAR_COMBO_CAT_DE_GRADO_PLANTILLA()
            Call CARGAR_COMBO_CAT_DE_GRADO_REAL()
            Call CARGAR_COMBO_CAT_DE_CATEGORIA_SALARIAL()
            Call CARGAR_COMBO_CAT_DE_SITUACION()
            Call CARGAR_COMBO_CAT_DE_ESTABLECIMIENTO()
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE CARGOS] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_C"
            Call BUSCAR_VISTA_MAESTRO_DE_CARGOS()
            Call BUSCAR_UBICACION()
            Call DESACTIVAR_TODO()
            Me.TextBox35.Text = vID_M_P
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 2 Then     'INFORMACION PERSONAL
            Call LIMPIAR_TEXTOS_2()
            Call CARGAR_COMBO_CAT_DE_TIPO_DE_SEXO()
            Call CARGAR_COMBO_CAT_DE_IPSS()
            Call CARGAR_COMBO_CAT_DE_TEZ()
            Call CARGAR_COMBO_CAT_DE_CABELLO()
            Call CARGAR_COMBO_CAT_DE_NIVEL_ACADEMICO()
            Call CARGAR_COMBO_CAT_DE_NIVEL_PROFESIONAL()
            Call CARGAR_COMBO_CAT_DE_NACIONALIDAD_N()
            Call CARGAR_COMBO_CAT_DE_DPTO_N()
            Call CARGAR_COMBO_CAT_DE_MUNICIPIO_N()
            Call CARGAR_COMBO_CAT_DE_NACIONALIDAD_D()
            Call CARGAR_COMBO_CAT_DE_DPTO_D()
            Call CARGAR_COMBO_CAT_DE_MUNICIPIO_D()
            Call CARGAR_COMBO_CAT_DE_ESTADO_CIVIL()
            Call CARGAR_COMBO_CAT_DE_TIPO_DE_HORARIO()
            Call CARGAR_COMBO_CAT_DE_TIPO_DE_AFILIACION()
            Call CARGAR_COMBO_CAT_DE_TIPO_DE_PLAZA()
            Call CARGAR_COMBO_CAT_DE_CATEGORIA_DE_LIC()
            Call CARGAR_COMBO_CAT_DE_TIPO_SANGUINEO()
            Call CARGAR_COMBO_CAT_DE_FUNCIONALIDAD()
            Call CARGAR_COMBO_CAT_DE_NIVEL_DE_COMPETENCIA()
            Call CARGAR_COMBO_CAT_DE_ADMINISTRADORES()
            If isuID_TIPO_USUARIO = 5 Then
                CADENAsql = "SELECT * FROM [VISTA MAESTRO DE PERSONAS] WHERE ID_M_P = " & vID_M_P & "  AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ID_M_P"
            Else
                CADENAsql = "SELECT * FROM [VISTA MAESTRO DE PERSONAS] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P"
            End If
            Call BUSCAR_VISTA_MAESTRO_DE_PERSONAS()
            Call BUSCAR_UBICACION()
            Call DESACTIVAR_TODO()
            Me.TextBox35.Text = vID_M_P
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 21 Then     'BANCOS
            If isuID_TIPO_USUARIO = 5 Then
                CADENAsql = "SELECT * FROM [VISTA MAESTRO DE PERSONAS] WHERE ID_M_P = " & vID_M_P & "  AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ID_M_P"
            Else
                CADENAsql = "SELECT * FROM [VISTA MAESTRO DE PERSONAS] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P"
            End If
            Call BUSCAR_VISTA_MAESTRO_DE_PERSONAS_BANCOS()
            Call BUSCAR_UBICACION()
            Call DESACTIVAR_TODO()
            Me.TextBox35.Text = vID_M_P
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 3 Then     'NÚCLEO FAMILIAR
            Me.DGV1.DataSource = Nothing
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE NUCLEO FAMILIAR] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_NF"
            Call CARGAR_DATAGRIDVIEW_1()
            Call ARMAR_DATAGRIDVIEW_1()
            Me.TextBox35.Text = vID_M_P
            Call DESACTIVAR_TODO()
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 4 Then     'EXPERIENCIA LABORAL
            Me.DGV2.DataSource = Nothing
            CADENAsql = "SELECT * FROM [MAESTRO DE EXPERIENCIA LABORAL] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_EL"
            Call CARGAR_DATAGRIDVIEW_2()
            Call ARMAR_DATAGRIDVIEW_2()
            Me.TextBox35.Text = vID_M_P
            Call DESACTIVAR_TODO()
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 5 Then     'REFERENCIAS PERSONALES
            Me.DGV3.DataSource = Nothing
            CADENAsql = "SELECT * FROM [MAESTRO DE REFERENCIAS PERSONALES] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_RP"
            Call CARGAR_DATAGRIDVIEW_3()
            Call ARMAR_DATAGRIDVIEW_3()
            Me.TextBox35.Text = vID_M_P
            Call DESACTIVAR_TODO()
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 6 Then     'ESTUDIOS
            Me.DGV4.DataSource = Nothing
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE ESTUDIOS] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_M_EST"
            Call CARGAR_DATAGRIDVIEW_4()
            Call ARMAR_DATAGRIDVIEW_4()
            Me.TextBox35.Text = vID_M_P
            Call DESACTIVAR_TODO()
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 7 Then     'HABILIDADES ADMINISTRATIVAS
            Me.DGV5.DataSource = Nothing
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HABILIDADES ADMIN] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_H_ADMIN"
            Call CARGAR_DATAGRIDVIEW_5()
            Call ARMAR_DATAGRIDVIEW_5()
            Me.TextBox35.Text = vID_M_P
            Call DESACTIVAR_TODO()
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 8 Then     'IDIOMAS
            Me.DGV6.DataSource = Nothing
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE IDIOMAS] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_M_I"
            Call CARGAR_DATAGRIDVIEW_6()
            Call ARMAR_DATAGRIDVIEW_6()
            Me.TextBox35.Text = vID_M_P
            Call DESACTIVAR_TODO()
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 9 Then     'CURVA Y TALLA
            Me.DGV7.DataSource = Nothing
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE CURVA Y TALLA] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_CT"
            Call CARGAR_DATAGRIDVIEW_7()
            Call ARMAR_DATAGRIDVIEW_7()
            Me.TextBox35.Text = vID_M_P
            Call DESACTIVAR_TODO()
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 10 Then    'CLASIFICACION DE PERSONAL
            Me.DGV8.DataSource = Nothing
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE CLASIFICACION DE PERSONAL] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_C_PERSONAL"
            Call CARGAR_DATAGRIDVIEW_8()
            Call ARMAR_DATAGRIDVIEW_8()
            Me.TextBox35.Text = vID_M_P
            Call DESACTIVAR_TODO()
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 11 Then    'CONDECORACIONES
            Me.DGV9.DataSource = Nothing
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE CONDECORACIONES] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_M_CONDE"
            Call CARGAR_DATAGRIDVIEW_9()
            Call ARMAR_DATAGRIDVIEW_9()
            Me.TextBox35.Text = vID_M_P
            Call DESACTIVAR_TODO()
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 12 Then    'EVENTUALIDADES
            Me.DGV10.DataSource = Nothing
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE EVENTUALIDADES] WHERE ID_M_P = " & vID_M_P & " ORDER BY FEC_EVE DESC"
            Call CARGAR_DATAGRIDVIEW_10()
            Call ARMAR_DATAGRIDVIEW_10()
            Me.TextBox35.Text = vID_M_P
            Call DESACTIVAR_TODO()
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 13 Then    'NOMINAS Y SALARIOS
            Me.DGV11.DataSource = Nothing
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE NOMINAS Y SALARIOS] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_NOMINA, F_MOVIMIENTO_N DESC"
            Call CARGAR_DATAGRIDVIEW_11()
            Call ARMAR_DATAGRIDVIEW_11()
            Me.TextBox35.Text = vID_M_P
            Call DESACTIVAR_TODO()
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 14 Then    'DOCUMENTOS DIGITALES
            Me.DGV12.DataSource = Nothing
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE DOCUMENTOS DIGITALES] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_T_DOC"
            Call CARGAR_DATAGRIDVIEW_12()
            Call ARMAR_DATAGRIDVIEW_12()
            Me.TextBox35.Text = vID_M_P
            Call DESACTIVAR_TODO()
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 15 Then    'SITUACION DE PERSONAL
            Me.DGV13.DataSource = Nothing
            RemoveHandler Me.DateTimePicker7.ValueChanged, AddressOf DateTimePicker7_ValueChanged
            RemoveHandler Me.DateTimePicker6.ValueChanged, AddressOf DateTimePicker6_ValueChanged
            '-------------*************----------------------------------------
            Call Clases.BUSCAR_FECHA_Y_HORA()
            Me.DateTimePicker7.Value = "01" & "/" & FECHA_SERVIDOR.Month & "/" & FECHA_SERVIDOR.Year
            Me.DateTimePicker6.Value = FECHA_SERVIDOR.Day & "/" & FECHA_SERVIDOR.Month & "/" & FECHA_SERVIDOR.Year
            FECHA01 = Me.DateTimePicker7.Text
            FECHA01 = Mid(FECHA01, 1, 10)
            FECHA02 = Me.DateTimePicker6.Text
            FECHA02 = Mid(FECHA02, 1, 10)
            '-------------*************----------------------------------------
            CADENAsql = "SET LANGUAGE ESPAÑOL Select ID, NIVEL1, NIVEL2, NIVEL3, NIVEL4, NIVEL5, ORD, ID_M_A, ID_M_P, CODIGO, UPPER(DATENAME (dw, DIA)) AS DIALETRAS, DIA, ID_T_SIT_P, D_T_SIT_P, ID_SIT_P, D_SIT_P, SIGLA, VALOR_SIT, VALOR_DIA, OBSERVACIONES, DETALLE_1, DETALLE_2, DETALLE_3 from [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE CODIGO = '" & Me.TextBox2.Text & "' AND DIA BETWEEN '" & FECHA01 & "' AND '" & FECHA02 & "' ORDER BY CODIGO, DIA SET LANGUAGE ENGLISH"
            Call CARGAR_DATAGRIDVIEW_13()
            Call ARMAR_DATAGRIDVIEW_13()
            Me.TextBox35.Text = vID_M_P
            Call DESACTIVAR_TODO()
            AddHandler Me.DateTimePicker7.ValueChanged, AddressOf DateTimePicker7_ValueChanged
            AddHandler Me.DateTimePicker6.ValueChanged, AddressOf DateTimePicker6_ValueChanged
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 16 Then    'HISTORICO DE MOVIMIENTOS
            Me.DGV14.DataSource = Nothing
            CADENAsql = "SELECT * FROM [VISTA HISTORICO DE MOVIMIENTO] WHERE ID_M_P = " & vID_M_P & " AND APLICAR = 1 ORDER BY ID_HIST_M DESC"
            Call CARGAR_DATAGRIDVIEW_14()
            Call ARMAR_DATAGRIDVIEW_14()
            Me.TextBox35.Text = vID_M_P
            Call DESACTIVAR_TODO()
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 17 Then    'MARCAS BIOMETRICAS
            Me.DGV15.DataSource = Nothing
            RemoveHandler Me.DateTimePicker4.ValueChanged, AddressOf DateTimePicker4_ValueChanged
            RemoveHandler Me.DateTimePicker5.ValueChanged, AddressOf DateTimePicker5_ValueChanged
            '-------------*************----------------------------------------
            Me.DateTimePicker4.Value = "01" & "/" & Now.Month & "/" & Now.Year
            Me.DateTimePicker5.Value = Now.Day & "/" & Now.Month & "/" & Now.Year
            FECHA01 = "Convert(VARCHAR(10), Convert(Date, '" + Me.DateTimePicker4.Value + "', 105), 23)"
            FECHA02 = "Convert(VARCHAR(10), Convert(Date, '" + Me.DateTimePicker5.Value + "', 105), 23)"
            '-------------*************----------------------------------------
            CADENAsql = "SET LANGUAGE Español SELECT [IDEMPRESA], [RELOJ], [ID_EMPLEADO], UPPER(DATENAME(dw, DIA)) AS DIAL, [FECHA_MARCA], [HORA_MARCA], [FECHA_CARGA], [ARCHIVO], [NUMERO], [PROGRAMADO], [OBSERVACION], [TIPO_MARCA] FROM [dbo].[VISTA MAESTRO DE MARCAS] WHERE (ID_EMPLEADO = '" & Me.TextBox2.Text & "') AND (FECHA_MARCA BETWEEN " & FECHA01 & " AND " & FECHA02 & ")  ORDER BY FECHA_MARCA, HORA_MARCA SET LANGUAGE ENGLISH"
            Call CARGAR_DATAGRIDVIEW_15()
            Call ARMAR_DATAGRIDVIEW_15()
            Me.TextBox35.Text = vID_M_P
            Call DESACTIVAR_TODO()
            AddHandler Me.DateTimePicker4.ValueChanged, AddressOf DateTimePicker4_ValueChanged
            AddHandler Me.DateTimePicker5.ValueChanged, AddressOf DateTimePicker5_ValueChanged
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 18 Then    'HORARIOS
            Me.DGV16.DataSource = Nothing
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HORARIOS] WHERE ID_M_P = " & vID_M_P & " ORDER BY HORA_E"
            Call CARGAR_DATAGRIDVIEW_16()
            Call ARMAR_DATAGRIDVIEW_16()
            Me.TextBox35.Text = vID_M_P
            Call DESACTIVAR_TODO()
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 19 Then    'SUBSIDIOS
            Me.DGV17.DataSource = Nothing
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE SUBSIDIOS] WHERE ID_M_P = " & vID_M_P & " ORDER BY FECHA_I"
            Call CARGAR_DATAGRIDVIEW_17()
            Call ARMAR_DATAGRIDVIEW_17()
            Me.TextBox35.Text = vID_M_P
            Call DESACTIVAR_TODO()
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 20 Then     'CONTRATOS Y EVALUACIONES
            Me.DGV18.DataSource = Nothing
            CADENAsql = "SELECT * FROM [dbo].[VISTA MAESTRO DE CONTRATOS Y EVALUACIONES] WHERE ID_M_P = " & vID_M_P & " ORDER BY [CODIGO], [ID_CONT]"
            Call CARGAR_DATAGRIDVIEW_18()
            Call ARMAR_DATAGRIDVIEW_18()
            Me.TextBox35.Text = vID_M_P
            Call DESACTIVAR_TODO()
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 22 Then     'MAESTRO DE INMUNIZACIONES
            Me.DGV19.DataSource = Nothing
            CADENAsql = "SELECT * FROM [dbo].[VISTA MAESTRO DE INMUNIZACIONES] WHERE ID_M_P = " & vID_M_P & " ORDER BY [ID_M_INM]"
            Call CARGAR_DATAGRIDVIEW_19()
            Call ARMAR_DATAGRIDVIEW_19()
            Me.TextBox35.Text = vID_M_P
            Call DESACTIVAR_TODO()
            GoTo SALIR
        End If
SALIR:
    End Sub
    Private Sub BUSCAR_VISTA_MAESTRO_DE_PERSONAS_BANCOS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                Me.Cursor = Cursors.WaitCursor
                'CUENTA DE BANCO
                If Not IsDBNull(MiDataTable.Rows(0).Item(53).ToString) Then : Me.TextBox26.Text = MiDataTable.Rows(0).Item(53).ToString : Else : Me.TextBox26.Text = "" : End If
                'NOMBRES Y APELLIDOS DEL BENEFICIARIO
                If Not IsDBNull(MiDataTable.Rows(0).Item(88).ToString) Then : Me.TextBox29.Text = MiDataTable.Rows(0).Item(88).ToString : Else : Me.TextBox29.Text = "" : End If
                'PAIS DE EMISION DE CEDULA
                If Not IsDBNull(MiDataTable.Rows(0).Item(89).ToString) Then : Me.TextBox43.Text = MiDataTable.Rows(0).Item(89).ToString : Else : Me.TextBox43.Text = "NICARAGUA" : End If
                If Val(MiDataTable.Rows(0).Item(90).ToString) <> 0 Then                  'FECHA DE EMISION DE CEDULA
                    Me.DateTimePicker8.Value = MiDataTable.Rows(0).Item(90).ToString
                    Me.DateTimePicker8.Checked = True
                    Me.DateTimePicker8.Enabled = False
                Else
                    Me.DateTimePicker8.Checked = False
                    Me.DateTimePicker8.Enabled = False
                End If
                If Val(MiDataTable.Rows(0).Item(91).ToString) <> 0 Then                  'FECHA DE VENCIMIENTO DE CEDULA
                    Me.DateTimePicker9.Value = MiDataTable.Rows(0).Item(91).ToString
                    Me.DateTimePicker9.Checked = True
                    Me.DateTimePicker9.Enabled = False
                Else
                    Me.DateTimePicker9.Checked = False
                    Me.DateTimePicker9.Enabled = False
                End If
                Me.Cursor = Cursors.Default
            End If
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
        DGV19.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE INMUNIZACIONES]")
        Me.Label116.Text = DGV19.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_19()
        Me.DGV19.Columns(0).HeaderText = "Id"
        Me.DGV19.Columns(0).Width = 30
        Me.DGV19.Columns(0).Visible = True
        Me.DGV19.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV19.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV19.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Me.DGV19.Columns(1).HeaderText = "ID_M_P"
        Me.DGV19.Columns(1).Width = 0
        Me.DGV19.Columns(1).Visible = False

        Me.DGV19.Columns(2).HeaderText = "ID_VACUNA"
        Me.DGV19.Columns(2).Width = 0
        Me.DGV19.Columns(2).Visible = False

        Me.DGV19.Columns(3).HeaderText = "Vacuna"
        Me.DGV19.Columns(3).Width = 200
        Me.DGV19.Columns(3).Visible = True
        Me.DGV19.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV19.Columns(4).HeaderText = "ID_DOSIS"
        Me.DGV19.Columns(4).Width = 0
        Me.DGV19.Columns(4).Visible = False

        Me.DGV19.Columns(5).HeaderText = "Dosis"
        Me.DGV19.Columns(5).Width = 80
        Me.DGV19.Columns(5).Visible = True
        Me.DGV19.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV19.Columns(6).HeaderText = "Fecha Vacuna"
        Me.DGV19.Columns(6).Width = 80
        Me.DGV19.Columns(6).Visible = True
        Me.DGV19.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV19.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_18()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE CONTRATOS Y EVALUACIONES]")
        DGV18.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE CONTRATOS Y EVALUACIONES]")
        Me.Label109.Text = DGV18.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_18()
        Me.DGV18.Columns(0).HeaderText = "Id"
        Me.DGV18.Columns(0).Width = 30
        Me.DGV18.Columns(0).Visible = True
        Me.DGV18.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV18.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV18.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Me.DGV18.Columns(1).HeaderText = "ID_M_P"
        Me.DGV18.Columns(1).Width = 0
        Me.DGV18.Columns(1).Visible = False

        Me.DGV18.Columns(2).HeaderText = "CODIGO"
        Me.DGV18.Columns(2).Width = 0
        Me.DGV18.Columns(2).Visible = False

        Me.DGV18.Columns(3).HeaderText = "FECHA_ING_PAME"
        Me.DGV18.Columns(3).Width = 0
        Me.DGV18.Columns(3).Visible = False

        Me.DGV18.Columns(4).HeaderText = "ID_CONT"
        Me.DGV18.Columns(4).Width = 0
        Me.DGV18.Columns(4).Visible = False

        Me.DGV18.Columns(5).HeaderText = "Tipo"
        Me.DGV18.Columns(5).Width = 200
        Me.DGV18.Columns(5).Visible = True
        Me.DGV18.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV18.Columns(6).HeaderText = "Desde"
        Me.DGV18.Columns(6).Width = 80
        Me.DGV18.Columns(6).Visible = True
        Me.DGV18.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV18.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV18.Columns(7).HeaderText = "Hasta"
        Me.DGV18.Columns(7).Width = 80
        Me.DGV18.Columns(7).Visible = True
        Me.DGV18.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV18.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV18.Columns(8).HeaderText = "Fir mado"
        Me.DGV18.Columns(8).Width = 50
        Me.DGV18.Columns(8).Visible = True
        Me.DGV18.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV18.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV18.Columns(9).HeaderText = "Fecha Evaluacion"
        Me.DGV18.Columns(9).Width = 80
        Me.DGV18.Columns(9).Visible = True
        Me.DGV18.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV18.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV18.Columns(10).HeaderText = "Calificacion Evaluacion"
        Me.DGV18.Columns(10).Width = 80
        Me.DGV18.Columns(10).Visible = True
        Me.DGV18.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV18.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        'Me.DGV18.Columns(10).DefaultCellStyle.BackColor = Color.Khaki

        Me.DGV18.Columns(11).HeaderText = ""
        Me.DGV18.Columns(11).Width = 0
        Me.DGV18.Columns(11).Visible = False

        Me.DGV18.Columns(12).HeaderText = "Clasificacion"
        Me.DGV18.Columns(12).Width = 100
        Me.DGV18.Columns(12).Visible = True
        Me.DGV18.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV18.Columns(13).HeaderText = "Observaciones"
        Me.DGV18.Columns(13).Width = 410
        Me.DGV18.Columns(13).Visible = True
        Me.DGV18.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_17()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE SUBSIDIOS]")
        DGV17.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE SUBSIDIOS]")
        Me.Label103.Text = DGV17.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_17()
        Me.DGV17.Columns(0).HeaderText = "Id"
        Me.DGV17.Columns(0).Width = 30
        Me.DGV17.Columns(0).Visible = True
        Me.DGV17.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV17.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV17.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Me.DGV17.Columns(1).HeaderText = "ID_M_P"
        Me.DGV17.Columns(1).Width = 0
        Me.DGV17.Columns(1).Visible = False

        Me.DGV17.Columns(2).HeaderText = "No. Exp."
        Me.DGV17.Columns(2).Width = 80
        Me.DGV17.Columns(2).Visible = True
        Me.DGV17.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV17.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV17.Columns(3).HeaderText = "Boleta"
        Me.DGV17.Columns(3).Width = 80
        Me.DGV17.Columns(3).Visible = True
        Me.DGV17.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV17.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV17.Columns(3).Frozen = True

        Me.DGV17.Columns(4).HeaderText = "ID_T_SUBS"
        Me.DGV17.Columns(4).Width = 0
        Me.DGV17.Columns(4).Visible = False

        Me.DGV17.Columns(5).HeaderText = "Tipo"
        Me.DGV17.Columns(5).Width = 220
        Me.DGV17.Columns(5).Visible = True
        Me.DGV17.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV17.Columns(6).HeaderText = "Ord"
        Me.DGV17.Columns(6).Width = 40
        Me.DGV17.Columns(6).Visible = True
        Me.DGV17.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV17.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Me.DGV17.Columns(7).HeaderText = "Fecha Inicio"
        Me.DGV17.Columns(7).Width = 90
        Me.DGV17.Columns(7).Visible = True
        Me.DGV17.Columns(7).DefaultCellStyle.Format = "d"
        Me.DGV17.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV17.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV17.Columns(8).HeaderText = "Fecha Fin"
        Me.DGV17.Columns(8).Width = 90
        Me.DGV17.Columns(8).Visible = True
        Me.DGV17.Columns(8).DefaultCellStyle.Format = "d"
        Me.DGV17.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV17.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV17.Columns(9).HeaderText = "Dias"
        Me.DGV17.Columns(9).Width = 40
        Me.DGV17.Columns(9).Visible = True
        Me.DGV17.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV17.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Me.DGV17.Columns(10).HeaderText = "Fecha Parto"
        Me.DGV17.Columns(10).Width = 90
        Me.DGV17.Columns(10).Visible = True
        Me.DGV17.Columns(10).DefaultCellStyle.Format = "d"
        Me.DGV17.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV17.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV17.Columns(11).HeaderText = "Fecha Parto Prob."
        Me.DGV17.Columns(11).Width = 90
        Me.DGV17.Columns(11).Visible = True
        Me.DGV17.Columns(11).DefaultCellStyle.Format = "d"
        Me.DGV17.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV17.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV17.Columns(12).HeaderText = "Fecha Acc. Lab."
        Me.DGV17.Columns(12).Width = 90
        Me.DGV17.Columns(12).Visible = True
        Me.DGV17.Columns(12).DefaultCellStyle.Format = "d"
        Me.DGV17.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV17.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV17.Columns(13).HeaderText = "Fecha Declaracion"
        Me.DGV17.Columns(13).Width = 90
        Me.DGV17.Columns(13).Visible = True
        Me.DGV17.Columns(13).DefaultCellStyle.Format = "d"
        Me.DGV17.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV17.Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV17.Columns(14).HeaderText = "VALOR_X_DIA"
        Me.DGV17.Columns(14).Width = 0
        Me.DGV17.Columns(14).Visible = False

        Me.DGV17.Columns(15).HeaderText = "VALOR_X_DIA"
        Me.DGV17.Columns(15).Width = 0
        Me.DGV17.Columns(15).Visible = False

        Me.DGV17.Columns(16).HeaderText = "Medico"
        Me.DGV17.Columns(16).Width = 230
        Me.DGV17.Columns(16).Visible = True
        Me.DGV17.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV17.Columns(17).HeaderText = "No. MINSA"
        Me.DGV17.Columns(17).Width = 60
        Me.DGV17.Columns(17).Visible = True
        Me.DGV17.Columns(17).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV17.Columns(17).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV17.Columns(18).HeaderText = "ID_E_MED"
        Me.DGV17.Columns(18).Width = 0
        Me.DGV17.Columns(18).Visible = False

        Me.DGV17.Columns(19).HeaderText = "Especialidad"
        Me.DGV17.Columns(19).Width = 0
        Me.DGV17.Columns(19).Visible = False
        Me.DGV17.Columns(19).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV17.Columns(20).HeaderText = "Diagnostico"
        Me.DGV17.Columns(20).Width = 500
        Me.DGV17.Columns(20).Visible = True
        Me.DGV17.Columns(20).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV17.Columns(21).HeaderText = "Observaciones"
        Me.DGV17.Columns(21).Width = 100
        Me.DGV17.Columns(21).Visible = True
        Me.DGV17.Columns(21).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_16()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE HORARIOS]")
        DGV16.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE HORARIOS]")
        Me.Label102.Text = DGV16.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_16()
        Me.DGV16.Columns(0).HeaderText = "Id"
        Me.DGV16.Columns(0).Width = 30
        Me.DGV16.Columns(0).Visible = True
        Me.DGV16.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV16.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV16.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Me.DGV16.Columns(1).HeaderText = "ID_M_P"
        Me.DGV16.Columns(1).Width = 0
        Me.DGV16.Columns(1).Visible = False

        Me.DGV16.Columns(2).HeaderText = "ID_T_HORARIO"
        Me.DGV16.Columns(2).Width = 0
        Me.DGV16.Columns(2).Visible = False

        Me.DGV16.Columns(3).HeaderText = "Horario"
        Me.DGV16.Columns(3).Width = 690
        Me.DGV16.Columns(3).Visible = True
        Me.DGV16.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV16.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV16.Columns(4).HeaderText = "Hora Entrada"
        Me.DGV16.Columns(4).Width = 120
        Me.DGV16.Columns(4).Visible = True
        Me.DGV16.Columns(4).DefaultCellStyle.Format = "t"
        Me.DGV16.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV16.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV16.Columns(5).HeaderText = "Hora Salida"
        Me.DGV16.Columns(5).Width = 120
        Me.DGV16.Columns(5).Visible = True
        Me.DGV16.Columns(5).DefaultCellStyle.Format = "t"
        Me.DGV16.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV16.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV16.Columns(6).HeaderText = "Horas Reales"
        Me.DGV16.Columns(6).Width = 150
        Me.DGV16.Columns(6).Visible = True
        Me.DGV16.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV16.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_15()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE MARCAS]")
        DGV15.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE MARCAS]")
        Me.Label97.Text = DGV15.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_15()
        Me.DGV15.Columns(0).HeaderText = "IDEMPRESA"
        Me.DGV15.Columns(0).Width = 0
        Me.DGV15.Columns(0).Visible = False

        Me.DGV15.Columns(1).HeaderText = "Reloj"
        Me.DGV15.Columns(1).Width = 330
        Me.DGV15.Columns(1).Visible = True
        Me.DGV15.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV15.Columns(2).HeaderText = "ID_EMPLEADO"
        Me.DGV15.Columns(2).Width = 0
        Me.DGV15.Columns(2).Visible = False
        Me.DGV15.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV15.Columns(3).HeaderText = "Dia"
        Me.DGV15.Columns(3).Width = 180
        Me.DGV15.Columns(3).Visible = True
        Me.DGV15.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV15.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        'Me.DGV15.Columns(3).DefaultCellStyle.BackColor = Color.Khaki

        Me.DGV15.Columns(4).HeaderText = "Fecha"
        Me.DGV15.Columns(4).Width = 180
        Me.DGV15.Columns(4).Visible = True
        Me.DGV15.Columns(4).DefaultCellStyle.Format = "d"
        Me.DGV15.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV15.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        'Me.DGV15.Columns(4).DefaultCellStyle.BackColor = Color.Khaki

        Me.DGV15.Columns(5).HeaderText = "Hora"
        Me.DGV15.Columns(5).Width = 180
        Me.DGV15.Columns(5).Visible = True
        Me.DGV15.Columns(5).DefaultCellStyle.Format = "t"
        Me.DGV15.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV15.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV15.Columns(6).HeaderText = "FECHA_CARGA"
        Me.DGV15.Columns(6).Width = 0
        Me.DGV15.Columns(6).Visible = False
        Me.DGV15.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV15.Columns(7).HeaderText = "ARCHIVO"
        Me.DGV15.Columns(7).Width = 0
        Me.DGV15.Columns(7).Visible = False
        Me.DGV15.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV15.Columns(8).HeaderText = "NUMERO"
        Me.DGV15.Columns(8).Width = 0
        Me.DGV15.Columns(8).Visible = False
        Me.DGV15.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV15.Columns(9).HeaderText = "PROGRAMADO"
        Me.DGV15.Columns(9).Width = 0
        Me.DGV15.Columns(9).Visible = False
        Me.DGV15.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV15.Columns(10).HeaderText = "OBSERVACION"
        Me.DGV15.Columns(10).Width = 0
        Me.DGV15.Columns(10).Visible = False
        Me.DGV15.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV15.Columns(11).HeaderText = "Tipo de Marca"
        Me.DGV15.Columns(11).Width = 220
        Me.DGV15.Columns(11).Visible = True
        Me.DGV15.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_14()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA HISTORICO DE MOVIMIENTO]")
        DGV14.DataSource = MiDataSet.Tables("[VISTA HISTORICO DE MOVIMIENTO]")
        Me.Label96.Text = DGV14.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_14()
        Me.DGV14.Columns(0).HeaderText = "Id"
        Me.DGV14.Columns(0).Width = 30
        Me.DGV14.Columns(0).Visible = True
        Me.DGV14.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV14.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV14.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Me.DGV14.Columns(1).HeaderText = "Fecha Ing. PAME"
        Me.DGV14.Columns(1).Width = 0
        Me.DGV14.Columns(1).Visible = False
        Me.DGV14.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV14.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV14.Columns(2).HeaderText = "ID_M_P"
        Me.DGV14.Columns(2).Width = 0
        Me.DGV14.Columns(2).Visible = False

        Me.DGV14.Columns(3).HeaderText = "Codigo"
        Me.DGV14.Columns(3).Width = 80
        Me.DGV14.Columns(3).Visible = True
        Me.DGV14.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV14.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV14.Columns(4).HeaderText = "Apellidos y Nombres"
        Me.DGV14.Columns(4).Width = 0
        Me.DGV14.Columns(4).Visible = False
        Me.DGV14.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV14.Columns(5).HeaderText = "Fecha Grabacion"
        Me.DGV14.Columns(5).Width = 80
        Me.DGV14.Columns(5).Visible = True
        Me.DGV14.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV14.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV14.Columns(6).HeaderText = "Fecha Aplicacion"
        Me.DGV14.Columns(6).Width = 80
        Me.DGV14.Columns(6).Visible = True
        Me.DGV14.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV14.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV14.Columns(7).HeaderText = "ID_T_M_EXP"
        Me.DGV14.Columns(7).Width = 0
        Me.DGV14.Columns(7).Visible = False

        Me.DGV14.Columns(8).HeaderText = "Tipo Movimiento"
        Me.DGV14.Columns(8).Width = 110
        Me.DGV14.Columns(8).Visible = True
        Me.DGV14.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV14.Columns(8).Frozen = True

        Me.DGV14.Columns(9).HeaderText = "ID_ST_M_L_EXP"
        Me.DGV14.Columns(9).Width = 0
        Me.DGV14.Columns(9).Visible = False

        Me.DGV14.Columns(10).HeaderText = "Causa Legal"
        Me.DGV14.Columns(10).Width = 160
        Me.DGV14.Columns(10).Visible = True
        Me.DGV14.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV14.Columns(11).HeaderText = "ID_ST_M_R_EXP"
        Me.DGV14.Columns(11).Width = 0
        Me.DGV14.Columns(11).Visible = False

        Me.DGV14.Columns(12).HeaderText = "Causa Real"
        Me.DGV14.Columns(12).Width = 160
        Me.DGV14.Columns(12).Visible = True
        Me.DGV14.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV14.Columns(13).HeaderText = "Cargo"
        Me.DGV14.Columns(13).Width = 160
        Me.DGV14.Columns(13).Visible = True
        Me.DGV14.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV14.Columns(14).HeaderText = "Gp"
        Me.DGV14.Columns(14).Width = 0
        Me.DGV14.Columns(14).Visible = False
        Me.DGV14.Columns(14).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV14.Columns(15).HeaderText = "Grado Real"
        Me.DGV14.Columns(15).Width = 160
        Me.DGV14.Columns(15).Visible = True
        Me.DGV14.Columns(15).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV14.Columns(16).HeaderText = "Nivel 1"
        Me.DGV14.Columns(16).Width = 60
        Me.DGV14.Columns(16).Visible = True
        Me.DGV14.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV14.Columns(17).HeaderText = "Nivel 2"
        Me.DGV14.Columns(17).Width = 60
        Me.DGV14.Columns(17).Visible = True
        Me.DGV14.Columns(17).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV14.Columns(18).HeaderText = "Nivel 3"
        Me.DGV14.Columns(18).Width = 60
        Me.DGV14.Columns(18).Visible = True
        Me.DGV14.Columns(18).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV14.Columns(19).HeaderText = "Nivel 4"
        Me.DGV14.Columns(19).Width = 60
        Me.DGV14.Columns(19).Visible = True
        Me.DGV14.Columns(19).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV14.Columns(20).HeaderText = "Nivel 5"
        Me.DGV14.Columns(20).Width = 60
        Me.DGV14.Columns(20).Visible = True
        Me.DGV14.Columns(20).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV14.Columns(21).HeaderText = "Nivel 6"
        Me.DGV14.Columns(21).Width = 60
        Me.DGV14.Columns(21).Visible = True
        Me.DGV14.Columns(21).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV14.Columns(22).HeaderText = "Nivel 7"
        Me.DGV14.Columns(22).Width = 60
        Me.DGV14.Columns(22).Visible = True
        Me.DGV14.Columns(22).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV14.Columns(23).HeaderText = "Nivel 8"
        Me.DGV14.Columns(23).Width = 60
        Me.DGV14.Columns(23).Visible = True
        Me.DGV14.Columns(23).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV14.Columns(24).HeaderText = "No. Orden"
        Me.DGV14.Columns(24).Width = 50
        Me.DGV14.Columns(24).Visible = True
        Me.DGV14.Columns(24).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV14.Columns(24).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV14.Columns(25).HeaderText = "Nomina"
        Me.DGV14.Columns(25).Width = 100
        Me.DGV14.Columns(25).Visible = True
        Me.DGV14.Columns(25).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV14.Columns(26).HeaderText = "Salario"
        Me.DGV14.Columns(26).Width = 70
        Me.DGV14.Columns(26).Visible = True
        Me.DGV14.Columns(26).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV14.Columns(26).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV14.Columns(26).DefaultCellStyle.Format = "##,##0.00"

        Me.DGV14.Columns(27).HeaderText = "Observaciones"
        Me.DGV14.Columns(27).Width = 150
        Me.DGV14.Columns(27).Visible = True
        Me.DGV14.Columns(27).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV14.Columns(28).HeaderText = "Mixta"
        Me.DGV14.Columns(28).Width = 0
        Me.DGV14.Columns(28).Visible = False
        Me.DGV14.Columns(28).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV14.Columns(29).HeaderText = "Militar"
        Me.DGV14.Columns(29).Width = 35
        Me.DGV14.Columns(29).Visible = True
        Me.DGV14.Columns(29).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV14.Columns(30).HeaderText = "PAME"
        Me.DGV14.Columns(30).Width = 35
        Me.DGV14.Columns(30).Visible = True
        Me.DGV14.Columns(30).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV14.Columns(31).HeaderText = "Apli cado"
        Me.DGV14.Columns(31).Width = 50
        Me.DGV14.Columns(31).Visible = True
        Me.DGV14.Columns(31).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV14.Columns(31).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV14.Columns(32).HeaderText = "Estadis tico"
        Me.DGV14.Columns(32).Width = 50
        Me.DGV14.Columns(32).Visible = True
        Me.DGV14.Columns(32).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV14.Columns(32).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV14.Columns(33).HeaderText = "Usuario que Actualiza"
        Me.DGV14.Columns(33).Width = 100
        Me.DGV14.Columns(33).Visible = True
        Me.DGV14.Columns(33).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV14.Columns(34).HeaderText = "Fecha de Actualizacion"
        Me.DGV14.Columns(34).Width = 80
        Me.DGV14.Columns(34).Visible = True
        Me.DGV14.Columns(34).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV14.Columns(34).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV14.Columns(35).HeaderText = "D_ADMIN"
        Me.DGV14.Columns(35).Width = 0
        Me.DGV14.Columns(35).Visible = False

    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_13()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE SITUACION DE PERSONAL]")
        DGV13.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE SITUACION DE PERSONAL]")
        Me.Label91.Text = DGV13.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_13()
        Me.DGV13.Columns(0).HeaderText = "Id"
        Me.DGV13.Columns(0).Width = 0
        Me.DGV13.Columns(0).Visible = False

        Me.DGV13.Columns(1).HeaderText = "NIVEL1"
        Me.DGV13.Columns(1).Width = 0
        Me.DGV13.Columns(1).Visible = False

        Me.DGV13.Columns(2).HeaderText = "NIVEL2"
        Me.DGV13.Columns(2).Width = 0
        Me.DGV13.Columns(2).Visible = False

        Me.DGV13.Columns(3).HeaderText = "NIVEL3"
        Me.DGV13.Columns(3).Width = 0
        Me.DGV13.Columns(3).Visible = False

        Me.DGV13.Columns(4).HeaderText = "NIVEL4"
        Me.DGV13.Columns(4).Width = 0
        Me.DGV13.Columns(4).Visible = False

        Me.DGV13.Columns(5).HeaderText = "NIVEL5"
        Me.DGV13.Columns(5).Width = 0
        Me.DGV13.Columns(5).Visible = False

        Me.DGV13.Columns(6).HeaderText = "ORD"
        Me.DGV13.Columns(6).Width = 0
        Me.DGV13.Columns(6).Visible = False

        Me.DGV13.Columns(7).HeaderText = "Id"
        Me.DGV13.Columns(7).Width = 20
        Me.DGV13.Columns(7).Visible = True
        Me.DGV13.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV13.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Me.DGV13.Columns(8).HeaderText = "ID_M_P"
        Me.DGV13.Columns(8).Width = 0
        Me.DGV13.Columns(8).Visible = False

        Me.DGV13.Columns(9).HeaderText = "Codigo"
        Me.DGV13.Columns(9).Width = 70
        Me.DGV13.Columns(9).Visible = True
        Me.DGV13.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV13.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV13.Columns(10).HeaderText = "Dia"
        Me.DGV13.Columns(10).Width = 70
        Me.DGV13.Columns(10).Visible = True
        Me.DGV13.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV13.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV13.Columns(10).DefaultCellStyle.BackColor = Color.Khaki

        Me.DGV13.Columns(11).HeaderText = "Fecha"
        Me.DGV13.Columns(11).Width = 70
        Me.DGV13.Columns(11).Visible = True
        Me.DGV13.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV13.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV13.Columns(11).DefaultCellStyle.BackColor = Color.Khaki

        Me.DGV13.Columns(12).HeaderText = "ID_T_SIT_P"
        Me.DGV13.Columns(12).Width = 0
        Me.DGV13.Columns(12).Visible = False

        Me.DGV13.Columns(13).HeaderText = "Tipo"
        Me.DGV13.Columns(13).Width = 140
        Me.DGV13.Columns(13).Visible = True
        Me.DGV13.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV13.Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV13.Columns(14).HeaderText = "ID_SIT_P"
        Me.DGV13.Columns(14).Width = 0
        Me.DGV13.Columns(14).Visible = False

        Me.DGV13.Columns(15).HeaderText = "Situacion"
        Me.DGV13.Columns(15).Width = 170
        Me.DGV13.Columns(15).Visible = True
        Me.DGV13.Columns(15).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV13.Columns(15).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV13.Columns(16).HeaderText = "SIGLA"
        Me.DGV13.Columns(16).Width = 0
        Me.DGV13.Columns(16).Visible = False

        Me.DGV13.Columns(17).HeaderText = "Valor Situacion"
        Me.DGV13.Columns(17).Width = 60
        Me.DGV13.Columns(17).Visible = True
        Me.DGV13.Columns(17).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV13.Columns(17).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Me.DGV13.Columns(18).HeaderText = "Valor Dia"
        Me.DGV13.Columns(18).Width = 110
        Me.DGV13.Columns(18).Visible = True
        Me.DGV13.Columns(18).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV13.Columns(18).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Me.DGV13.Columns(19).HeaderText = "Observaciones"
        Me.DGV13.Columns(19).Width = 380
        Me.DGV13.Columns(19).Visible = True
        Me.DGV13.Columns(19).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV13.Columns(19).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV13.Columns(20).HeaderText = "DETALLE 1"
        Me.DGV13.Columns(20).Width = 0
        Me.DGV13.Columns(20).Visible = False

        Me.DGV13.Columns(21).HeaderText = "DETALLE 2"
        Me.DGV13.Columns(21).Width = 0
        Me.DGV13.Columns(21).Visible = False

        Me.DGV13.Columns(22).HeaderText = "DETALLE 3"
        Me.DGV13.Columns(22).Width = 0
        Me.DGV13.Columns(22).Visible = False
    End Sub
    Private Sub DGV13_RowPrePaint(sender As Object, e As DataGridViewRowPrePaintEventArgs) Handles DGV13.RowPrePaint
        Dim style As New DataGridViewCellStyle
        style.Font = New Font(DGV13.Font, FontStyle.Bold)
        For Each Row As DataGridViewRow In DGV13.Rows
            If Row.Cells(15).Value = "AUSENCIA INJUSTIFICADA" Or Row.Cells(15).Value = "OMISION DE MARCA" Then
                Row.DefaultCellStyle.ForeColor = Color.Red
            End If
            If Row.Cells(15).Value = "VACACIONES" Then
                Row.DefaultCellStyle.ForeColor = Color.Blue
            End If
            If Row.Cells(15).Value = "SUBSIDIO" Then
                Row.DefaultCellStyle.ForeColor = Color.DarkOrange
            End If
        Next
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_12()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE DOCUMENTOS DIGITALES]")
        DGV12.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE DOCUMENTOS DIGITALES]")
        Me.Label90.Text = DGV12.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_12()
        Me.DGV12.Columns(0).HeaderText = "Id"
        Me.DGV12.Columns(0).Width = 30
        Me.DGV12.Columns(0).Visible = True
        Me.DGV12.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV12.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV12.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Me.DGV12.Columns(1).HeaderText = "ID_M_P"
        Me.DGV12.Columns(1).Width = 0
        Me.DGV12.Columns(1).Visible = False

        Me.DGV12.Columns(2).HeaderText = "ID_D1"
        Me.DGV12.Columns(2).Width = 0
        Me.DGV12.Columns(2).Visible = False

        Me.DGV12.Columns(3).HeaderText = "DIRECTORIO"
        Me.DGV12.Columns(3).Width = 0
        Me.DGV12.Columns(3).Visible = False

        Me.DGV12.Columns(4).HeaderText = "ID_T_DOC"
        Me.DGV12.Columns(4).Width = 0
        Me.DGV12.Columns(4).Visible = False

        Me.DGV12.Columns(5).HeaderText = "Tipo de Documento"
        Me.DGV12.Columns(5).Width = 230
        Me.DGV12.Columns(5).Visible = True
        Me.DGV12.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV12.Columns(6).HeaderText = "Nombre"
        Me.DGV12.Columns(6).Width = 350
        Me.DGV12.Columns(6).Visible = True
        Me.DGV12.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV12.Columns(7).HeaderText = "Observacion"
        Me.DGV12.Columns(7).Width = 500
        Me.DGV12.Columns(7).Visible = True
        Me.DGV12.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_11()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE NOMINAS Y SALARIOS]")
        DGV11.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE NOMINAS Y SALARIOS]")
        Me.Label89.Text = DGV11.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_11()
        Me.DGV11.Columns(0).HeaderText = "Id"
        Me.DGV11.Columns(0).Width = 30
        Me.DGV11.Columns(0).Visible = True
        Me.DGV11.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV11.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV11.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Me.DGV11.Columns(1).HeaderText = "ID_M_P"
        Me.DGV11.Columns(1).Width = 0
        Me.DGV11.Columns(1).Visible = False

        Me.DGV11.Columns(2).HeaderText = "ID_MOVIMIENTO_N"
        Me.DGV11.Columns(2).Width = 0
        Me.DGV11.Columns(2).Visible = False

        Me.DGV11.Columns(3).HeaderText = "Movimiento Nómina"
        Me.DGV11.Columns(3).Width = 150
        Me.DGV11.Columns(3).Visible = True
        Me.DGV11.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV11.Columns(4).HeaderText = "Fecha"
        Me.DGV11.Columns(4).Width = 70
        Me.DGV11.Columns(4).Visible = True
        Me.DGV11.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV11.Columns(5).HeaderText = "ID_NOMINA"
        Me.DGV11.Columns(5).Width = 0
        Me.DGV11.Columns(5).Visible = False

        Me.DGV11.Columns(6).HeaderText = "Nomina"
        Me.DGV11.Columns(6).Width = 210
        Me.DGV11.Columns(6).Visible = True
        Me.DGV11.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV11.Columns(7).HeaderText = "Propuesta"
        Me.DGV11.Columns(7).Width = 70
        Me.DGV11.Columns(7).Visible = True
        Me.DGV11.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV11.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV11.Columns(7).DefaultCellStyle.Format = "##,##0.00"

        Me.DGV11.Columns(8).HeaderText = "Salario"
        Me.DGV11.Columns(8).Width = 70
        Me.DGV11.Columns(8).Visible = True
        Me.DGV11.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV11.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV11.Columns(8).DefaultCellStyle.Format = "##,##0.00"

        Me.DGV11.Columns(9).HeaderText = "Salario Minimo"
        Me.DGV11.Columns(9).Width = 60
        Me.DGV11.Columns(9).Visible = True
        Me.DGV11.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV11.Columns(10).HeaderText = "Observaciones"
        Me.DGV11.Columns(10).Width = 400
        Me.DGV11.Columns(10).Visible = True
        Me.DGV11.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV11.Columns(11).HeaderText = "Activo"
        Me.DGV11.Columns(11).Width = 50
        Me.DGV11.Columns(11).Visible = True
        Me.DGV11.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_10()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE EVENTUALIDADES]")
        DGV10.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE EVENTUALIDADES]")
        Me.Label88.Text = DGV10.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_10()
        Me.DGV10.Columns(0).HeaderText = "Id"
        Me.DGV10.Columns(0).Width = 30
        Me.DGV10.Columns(0).Visible = True
        Me.DGV10.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV10.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV10.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Me.DGV10.Columns(1).HeaderText = "ID_M_P"
        Me.DGV10.Columns(1).Width = 0
        Me.DGV10.Columns(1).Visible = False

        Me.DGV10.Columns(2).HeaderText = "Nivel 1"
        Me.DGV10.Columns(2).Width = 0
        Me.DGV10.Columns(2).Visible = False
        Me.DGV10.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV10.Columns(3).HeaderText = "Nivel 2"
        Me.DGV10.Columns(3).Width = 0
        Me.DGV10.Columns(3).Visible = 0
        Me.DGV10.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV10.Columns(4).HeaderText = "Jefe Inmediato"
        Me.DGV10.Columns(4).Width = 100
        Me.DGV10.Columns(4).Visible = True
        Me.DGV10.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV10.Columns(5).HeaderText = "Fecha"
        Me.DGV10.Columns(5).Width = 70
        Me.DGV10.Columns(5).Visible = True
        Me.DGV10.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV10.Columns(6).HeaderText = "ID_CR_EVENTOS"
        Me.DGV10.Columns(6).Width = 0
        Me.DGV10.Columns(6).Visible = False

        Me.DGV10.Columns(7).HeaderText = "Causa"
        Me.DGV10.Columns(7).Width = 100
        Me.DGV10.Columns(7).Visible = True
        Me.DGV10.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV10.Columns(8).HeaderText = "ID_E"
        Me.DGV10.Columns(8).Width = 0
        Me.DGV10.Columns(8).Visible = False

        Me.DGV10.Columns(9).HeaderText = "Eventualidad"
        Me.DGV10.Columns(9).Width = 140
        Me.DGV10.Columns(9).Visible = True
        Me.DGV10.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV10.Columns(10).HeaderText = "Observaciones"
        Me.DGV10.Columns(10).Width = 670
        Me.DGV10.Columns(10).Visible = True
        Me.DGV10.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_9()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE CONDECORACIONES]")
        DGV9.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE CONDECORACIONES]")
        Me.Label87.Text = DGV9.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_9()
        Me.DGV9.Columns(0).HeaderText = "ID_C_CONDE"
        Me.DGV9.Columns(0).Width = 0
        Me.DGV9.Columns(0).Visible = False

        Me.DGV9.Columns(1).HeaderText = "Clasificacion"
        Me.DGV9.Columns(1).Width = 250
        Me.DGV9.Columns(1).Visible = True
        Me.DGV9.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV9.Columns(2).HeaderText = "ID_CONDE"
        Me.DGV9.Columns(2).Width = 0
        Me.DGV9.Columns(2).Visible = False

        Me.DGV9.Columns(3).HeaderText = "Codigo"
        Me.DGV9.Columns(3).Width = 80
        Me.DGV9.Columns(3).Visible = True
        Me.DGV9.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV9.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        'Me.DGV9.Columns(3).Frozen = True

        Me.DGV9.Columns(4).HeaderText = "Condecoracion"
        Me.DGV9.Columns(4).Width = 320
        Me.DGV9.Columns(4).Visible = True
        Me.DGV9.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV9.Columns(5).HeaderText = "Id"
        Me.DGV9.Columns(5).Width = 30
        Me.DGV9.Columns(5).Visible = True
        Me.DGV9.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV9.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV9.Columns(5).DefaultCellStyle.BackColor = Color.Black

        Me.DGV9.Columns(6).HeaderText = "ID_M_P"
        Me.DGV9.Columns(6).Width = 0
        Me.DGV9.Columns(6).Visible = False

        Me.DGV9.Columns(7).HeaderText = "Año de Entrega"
        Me.DGV9.Columns(7).Width = 100
        Me.DGV9.Columns(7).Visible = True
        Me.DGV9.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV9.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV9.Columns(8).HeaderText = "Motivos"
        Me.DGV9.Columns(8).Width = 120
        Me.DGV9.Columns(8).Visible = True
        Me.DGV9.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV9.Columns(9).HeaderText = "Observaciones"
        Me.DGV9.Columns(9).Width = 200
        Me.DGV9.Columns(9).Visible = True
        Me.DGV9.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_8()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE CLASIFICACION DE PERSONAL]")
        DGV8.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE CLASIFICACION DE PERSONAL]")
        Me.Label86.Text = DGV8.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_8()
        Me.DGV8.Columns(0).HeaderText = "Id"
        Me.DGV8.Columns(0).Width = 30
        Me.DGV8.Columns(0).Visible = True
        Me.DGV8.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV8.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV8.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Me.DGV8.Columns(1).HeaderText = "ID_M_P"
        Me.DGV8.Columns(1).Width = 0
        Me.DGV8.Columns(1).Visible = False

        Me.DGV8.Columns(2).HeaderText = "ID_C_PE"
        Me.DGV8.Columns(2).Width = 0
        Me.DGV8.Columns(2).Visible = False

        Me.DGV8.Columns(3).HeaderText = "Clasificacion"
        Me.DGV8.Columns(3).Width = 210
        Me.DGV8.Columns(3).Visible = True
        Me.DGV8.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV8.Columns(4).HeaderText = "Observaciones"
        Me.DGV8.Columns(4).Width = 870
        Me.DGV8.Columns(4).Visible = True
        Me.DGV8.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_7()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE CURVA Y TALLA]")
        DGV7.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE CURVA Y TALLA]")
        Me.Label85.Text = DGV7.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_7()
        Me.DGV7.Columns(0).HeaderText = "Id"
        Me.DGV7.Columns(0).Width = 30
        Me.DGV7.Columns(0).Visible = True
        Me.DGV7.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV7.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV7.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Me.DGV7.Columns(1).HeaderText = "ID_M_P"
        Me.DGV7.Columns(1).Width = 0
        Me.DGV7.Columns(1).Visible = False

        Me.DGV7.Columns(2).HeaderText = "ID_T_CT"
        Me.DGV7.Columns(2).Width = 0
        Me.DGV7.Columns(2).Visible = False

        Me.DGV7.Columns(3).HeaderText = "Tipo"
        Me.DGV7.Columns(3).Width = 200
        Me.DGV7.Columns(3).Visible = True
        Me.DGV7.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV7.Columns(4).HeaderText = "Talla"
        Me.DGV7.Columns(4).Width = 60
        Me.DGV7.Columns(4).Visible = True
        Me.DGV7.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV7.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV7.Columns(5).HeaderText = "Observaciones"
        Me.DGV7.Columns(5).Width = 820
        Me.DGV7.Columns(5).Visible = True
        Me.DGV7.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_6()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE IDIOMAS]")
        DGV6.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE IDIOMAS]")
        Me.Label84.Text = DGV6.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_6()
        Me.DGV6.Columns(0).HeaderText = "Id"
        Me.DGV6.Columns(0).Width = 30
        Me.DGV6.Columns(0).Visible = True
        Me.DGV6.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV6.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV6.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Me.DGV6.Columns(1).HeaderText = "ID_M_P"
        Me.DGV6.Columns(1).Width = 0
        Me.DGV6.Columns(1).Visible = False

        Me.DGV6.Columns(2).HeaderText = "ID_IDIOMA"
        Me.DGV6.Columns(2).Width = 0
        Me.DGV6.Columns(2).Visible = False

        Me.DGV6.Columns(3).HeaderText = "Idioma"
        Me.DGV6.Columns(3).Width = 160
        Me.DGV6.Columns(3).Visible = True
        Me.DGV6.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV6.Columns(4).HeaderText = "Habla"
        Me.DGV6.Columns(4).Width = 70
        Me.DGV6.Columns(4).Visible = True
        Me.DGV6.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV6.Columns(5).HeaderText = "%"
        Me.DGV6.Columns(5).Width = 70
        Me.DGV6.Columns(5).Visible = True
        Me.DGV6.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV6.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV6.Columns(6).HeaderText = "Lee"
        Me.DGV6.Columns(6).Width = 70
        Me.DGV6.Columns(6).Visible = True
        Me.DGV6.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV6.Columns(7).HeaderText = "%"
        Me.DGV6.Columns(7).Width = 70
        Me.DGV6.Columns(7).Visible = True
        Me.DGV6.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV6.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV6.Columns(8).HeaderText = "Escribe"
        Me.DGV6.Columns(8).Width = 70
        Me.DGV6.Columns(8).Visible = True
        Me.DGV6.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV6.Columns(9).HeaderText = "%"
        Me.DGV6.Columns(9).Width = 70
        Me.DGV6.Columns(9).Visible = True
        Me.DGV6.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV6.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV6.Columns(10).HeaderText = "Observaciones"
        Me.DGV6.Columns(10).Width = 500
        Me.DGV6.Columns(10).Visible = True
        Me.DGV6.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_5()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE HABILIDADES ADMIN]")
        DGV5.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE HABILIDADES ADMIN]")
        Me.Label83.Text = DGV5.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_5()
        Me.DGV5.Columns(0).HeaderText = "Id"
        Me.DGV5.Columns(0).Width = 30
        Me.DGV5.Columns(0).Visible = True
        Me.DGV5.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV5.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV5.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Me.DGV5.Columns(1).HeaderText = "ID_M_P"
        Me.DGV5.Columns(1).Width = 0
        Me.DGV5.Columns(1).Visible = False

        Me.DGV5.Columns(2).HeaderText = "ID_HABILIDAD"
        Me.DGV5.Columns(2).Width = 0
        Me.DGV5.Columns(2).Visible = False

        Me.DGV5.Columns(3).HeaderText = "Habilidad"
        Me.DGV5.Columns(3).Width = 200
        Me.DGV5.Columns(3).Visible = True
        Me.DGV5.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV5.Columns(4).HeaderText = "Porcentaje (%)"
        Me.DGV5.Columns(4).Width = 80
        Me.DGV5.Columns(4).Visible = True
        Me.DGV5.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV5.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV5.Columns(5).HeaderText = "Observaciones"
        Me.DGV5.Columns(5).Width = 800
        Me.DGV5.Columns(5).Visible = True
        Me.DGV5.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_4()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE ESTUDIOS]")
        DGV4.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE ESTUDIOS]")
        Me.Label82.Text = DGV4.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_4()
        Me.DGV4.Columns(0).HeaderText = "Id"
        Me.DGV4.Columns(0).Width = 30
        Me.DGV4.Columns(0).Visible = True
        Me.DGV4.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV4.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV4.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Me.DGV4.Columns(1).HeaderText = "ID_M_P"
        Me.DGV4.Columns(1).Width = 0
        Me.DGV4.Columns(1).Visible = False

        Me.DGV4.Columns(2).HeaderText = "ID_T_ESTUDIO"
        Me.DGV4.Columns(2).Width = 0
        Me.DGV4.Columns(2).Visible = False

        Me.DGV4.Columns(3).HeaderText = "Tipo de Estudio"
        Me.DGV4.Columns(3).Width = 100
        Me.DGV4.Columns(3).Visible = True
        Me.DGV4.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV4.Columns(3).Frozen = True

        Me.DGV4.Columns(4).HeaderText = "ID_E_MED"
        Me.DGV4.Columns(4).Width = 0
        Me.DGV4.Columns(4).Visible = False

        Me.DGV4.Columns(5).HeaderText = "Especialidad"
        Me.DGV4.Columns(5).Width = 150
        Me.DGV4.Columns(5).Visible = True
        Me.DGV4.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV4.Columns(6).HeaderText = "Ult. Grado"
        Me.DGV4.Columns(6).Width = 50
        Me.DGV4.Columns(6).Visible = True
        Me.DGV4.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV4.Columns(7).HeaderText = "Fecha de Estudio"
        Me.DGV4.Columns(7).Width = 80
        Me.DGV4.Columns(7).Visible = True
        Me.DGV4.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV4.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV4.Columns(8).HeaderText = "Nombre del Estudio"
        Me.DGV4.Columns(8).Width = 280
        Me.DGV4.Columns(8).Visible = True
        Me.DGV4.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV4.Columns(9).HeaderText = "Nombre Centro del Estudio"
        Me.DGV4.Columns(9).Width = 250
        Me.DGV4.Columns(9).Visible = True
        Me.DGV4.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV4.Columns(10).HeaderText = "Direccion Centro del Estudio"
        Me.DGV4.Columns(10).Width = 200
        Me.DGV4.Columns(10).Visible = True
        Me.DGV4.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV4.Columns(11).HeaderText = "Estudio Extranjero"
        Me.DGV4.Columns(11).Width = 70
        Me.DGV4.Columns(11).Visible = True
        Me.DGV4.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV4.Columns(12).HeaderText = "ID_PAIS"
        Me.DGV4.Columns(12).Width = 0
        Me.DGV4.Columns(12).Visible = False

        Me.DGV4.Columns(13).HeaderText = "Pais"
        Me.DGV4.Columns(13).Width = 100
        Me.DGV4.Columns(13).Visible = True
        Me.DGV4.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV4.Columns(14).HeaderText = "Titulo Obtenido"
        Me.DGV4.Columns(14).Width = 280
        Me.DGV4.Columns(14).Visible = True
        Me.DGV4.Columns(14).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV4.Columns(15).HeaderText = "Observaciones"
        Me.DGV4.Columns(15).Width = 280
        Me.DGV4.Columns(15).Visible = True
        Me.DGV4.Columns(15).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_3()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[MAESTRO DE REFERENCIAS PERSONALES]")
        DGV3.DataSource = MiDataSet.Tables("[MAESTRO DE REFERENCIAS PERSONALES]")
        Me.Label81.Text = DGV3.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_3()
        Me.DGV3.Columns(0).HeaderText = "Id"
        Me.DGV3.Columns(0).Width = 30
        Me.DGV3.Columns(0).Visible = True
        Me.DGV3.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV3.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV3.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Me.DGV3.Columns(1).HeaderText = "ID_M_P"
        Me.DGV3.Columns(1).Width = 0
        Me.DGV3.Columns(1).Visible = False

        Me.DGV3.Columns(2).HeaderText = "Apellidos y Nombres"
        Me.DGV3.Columns(2).Width = 310
        Me.DGV3.Columns(2).Visible = True
        Me.DGV3.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(3).HeaderText = "Direccion"
        Me.DGV3.Columns(3).Width = 420
        Me.DGV3.Columns(3).Visible = True
        Me.DGV3.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(4).HeaderText = "Ocupacion"
        Me.DGV3.Columns(4).Width = 250
        Me.DGV3.Columns(4).Visible = True
        Me.DGV3.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(5).HeaderText = "Telefono"
        Me.DGV3.Columns(5).Width = 100
        Me.DGV3.Columns(5).Visible = True
        Me.DGV3.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_2()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[MAESTRO DE EXPERIENCIA LABORAL]")
        DGV2.DataSource = MiDataSet.Tables("[MAESTRO DE EXPERIENCIA LABORAL]")
        Me.Label80.Text = DGV2.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_2()
        Me.DGV2.Columns(0).HeaderText = "Id"
        Me.DGV2.Columns(0).Width = 30
        Me.DGV2.Columns(0).Visible = True
        Me.DGV2.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV2.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV2.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Me.DGV2.Columns(1).HeaderText = "ID_M_P"
        Me.DGV2.Columns(1).Width = 0
        Me.DGV2.Columns(1).Visible = False

        Me.DGV2.Columns(2).HeaderText = "Nombre de Empresa"
        Me.DGV2.Columns(2).Width = 200
        Me.DGV2.Columns(2).Visible = True
        Me.DGV2.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV2.Columns(2).Frozen = True

        Me.DGV2.Columns(3).HeaderText = "Direccion de Empresa"
        Me.DGV2.Columns(3).Width = 200
        Me.DGV2.Columns(3).Visible = True
        Me.DGV2.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV2.Columns(4).HeaderText = "Telefono de Empresa"
        Me.DGV2.Columns(4).Width = 110
        Me.DGV2.Columns(4).Visible = True
        Me.DGV2.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV2.Columns(5).HeaderText = "Actividad de Empresa"
        Me.DGV2.Columns(5).Width = 200
        Me.DGV2.Columns(5).Visible = True
        Me.DGV2.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV2.Columns(6).HeaderText = "Cargo que Desempeñado"
        Me.DGV2.Columns(6).Width = 200
        Me.DGV2.Columns(6).Visible = True
        Me.DGV2.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV2.Columns(7).HeaderText = "Jefe Inmediato"
        Me.DGV2.Columns(7).Width = 200
        Me.DGV2.Columns(7).Visible = True
        Me.DGV2.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV2.Columns(8).HeaderText = "Salario"
        Me.DGV2.Columns(8).Width = 70
        Me.DGV2.Columns(8).Visible = True
        Me.DGV2.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV2.Columns(8).DefaultCellStyle.Format = "##,##0.00"

        Me.DGV2.Columns(9).HeaderText = "Desde"
        Me.DGV2.Columns(9).Width = 70
        Me.DGV2.Columns(9).Visible = True
        Me.DGV2.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV2.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Me.DGV2.Columns(10).HeaderText = "Hasta"
        Me.DGV2.Columns(10).Width = 70
        Me.DGV2.Columns(10).Visible = True
        Me.DGV2.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV2.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Me.DGV2.Columns(11).HeaderText = "Causa de Retiro"
        Me.DGV2.Columns(11).Width = 200
        Me.DGV2.Columns(11).Visible = True
        Me.DGV2.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV2.Columns(12).HeaderText = "Obligaciones y Responsabildiades"
        Me.DGV2.Columns(12).Width = 200
        Me.DGV2.Columns(12).Visible = True
        Me.DGV2.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_1()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE NUCLEO FAMILIAR]")
        DGV1.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE NUCLEO FAMILIAR]")
        Me.Label61.Text = DGV1.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_1()
        Me.DGV1.Columns(0).HeaderText = "Id"
        Me.DGV1.Columns(0).Width = 30
        Me.DGV1.Columns(0).Visible = True
        Me.DGV1.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV1.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Me.DGV1.Columns(1).HeaderText = "ID_M_P"
        Me.DGV1.Columns(1).Width = 0
        Me.DGV1.Columns(1).Visible = False

        Me.DGV1.Columns(2).HeaderText = "ID_PARENTELA"
        Me.DGV1.Columns(2).Width = 0
        Me.DGV1.Columns(2).Visible = False

        Me.DGV1.Columns(3).HeaderText = "Parentela"
        Me.DGV1.Columns(3).Width = 60
        Me.DGV1.Columns(3).Visible = True
        Me.DGV1.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(4).HeaderText = "Apellidos y Nombres"
        Me.DGV1.Columns(4).Width = 230
        Me.DGV1.Columns(4).Visible = True
        Me.DGV1.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV1.Columns(4).Frozen = True

        Me.DGV1.Columns(5).HeaderText = "ID_T_SEXO"
        Me.DGV1.Columns(5).Width = 0
        Me.DGV1.Columns(5).Visible = False

        Me.DGV1.Columns(6).HeaderText = "Sexo"
        Me.DGV1.Columns(6).Width = 70
        Me.DGV1.Columns(6).Visible = True
        Me.DGV1.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(7).HeaderText = "Fecha Nacimiento"
        Me.DGV1.Columns(7).Width = 70
        Me.DGV1.Columns(7).Visible = True
        Me.DGV1.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(8).HeaderText = "Edad"
        Me.DGV1.Columns(8).Width = 30
        Me.DGV1.Columns(8).Visible = True
        Me.DGV1.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV1.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Me.DGV1.Columns(9).HeaderText = "ID_N_ACADEMICO"
        Me.DGV1.Columns(9).Width = 0
        Me.DGV1.Columns(9).Visible = False

        Me.DGV1.Columns(10).HeaderText = "Nivel Academico"
        Me.DGV1.Columns(10).Width = 80
        Me.DGV1.Columns(10).Visible = True
        Me.DGV1.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(11).HeaderText = "Dirección Domiciliar"
        Me.DGV1.Columns(11).Width = 250
        Me.DGV1.Columns(11).Visible = True
        Me.DGV1.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(12).HeaderText = "Telefono 1"
        Me.DGV1.Columns(12).Width = 60
        Me.DGV1.Columns(12).Visible = True
        Me.DGV1.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(13).HeaderText = "Telefono 2"
        Me.DGV1.Columns(13).Width = 60
        Me.DGV1.Columns(13).Visible = True
        Me.DGV1.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(14).HeaderText = "Lugar de Trabajo"
        Me.DGV1.Columns(14).Width = 150
        Me.DGV1.Columns(14).Visible = True
        Me.DGV1.Columns(14).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(15).HeaderText = "Direccion de Trabajo"
        Me.DGV1.Columns(15).Width = 150
        Me.DGV1.Columns(15).Visible = True
        Me.DGV1.Columns(15).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(16).HeaderText = "Cargo que Ocupa"
        Me.DGV1.Columns(16).Width = 150
        Me.DGV1.Columns(16).Visible = True
        Me.DGV1.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(17).HeaderText = "Fallecido"
        Me.DGV1.Columns(17).Width = 50
        Me.DGV1.Columns(17).Visible = True
        Me.DGV1.Columns(17).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub BUSCAR_VISTA_MAESTRO_DE_CARGOS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                Me.Cursor = Cursors.WaitCursor
                Me.TextBox4.Text = MiDataTable.Rows(0).Item(28).ToString            'ORDEN MIXTO
                Me.TextBox37.Text = MiDataTable.Rows(0).Item(29).ToString           'ORDEN MILITAR
                Me.TextBox38.Text = MiDataTable.Rows(0).Item(30).ToString           'ORDEN PAME
                'Me.ComboBox1.Text = MiDataTable.Rows(0).Item(4).ToString            'N1
                'Me.ComboBox2.Text = MiDataTable.Rows(0).Item(7).ToString            'N2 
                'Me.ComboBox29.Text = MiDataTable.Rows(0).Item(10).ToString          'N3 
                'Me.ComboBox30.Text = MiDataTable.Rows(0).Item(13).ToString          'N4
                'Me.ComboBox31.Text = MiDataTable.Rows(0).Item(16).ToString          'N5
                'Me.ComboBox34.Text = MiDataTable.Rows(0).Item(19).ToString          'N6
                'Me.ComboBox35.Text = MiDataTable.Rows(0).Item(22).ToString          'N7
                'Me.ComboBox36.Text = MiDataTable.Rows(0).Item(25).ToString          'N8
                If Not IsDBNull(MiDataTable.Rows(0).Item(7).ToString) Then : Me.ComboBox2.Text = MiDataTable.Rows(0).Item(7).ToString : Else : Me.ComboBox2.Text = "" : End If
                If Not IsDBNull(MiDataTable.Rows(0).Item(10).ToString) Then : Me.ComboBox29.Text = MiDataTable.Rows(0).Item(10).ToString : Else : Me.ComboBox29.Text = "" : End If
                If Not IsDBNull(MiDataTable.Rows(0).Item(13).ToString) Then : Me.ComboBox30.Text = MiDataTable.Rows(0).Item(13).ToString : Else : Me.ComboBox30.Text = "" : End If
                If Not IsDBNull(MiDataTable.Rows(0).Item(16).ToString) Then : Me.ComboBox31.Text = MiDataTable.Rows(0).Item(16).ToString : Else : Me.ComboBox31.Text = "" : End If
                If Not IsDBNull(MiDataTable.Rows(0).Item(19).ToString) Then : Me.ComboBox34.Text = MiDataTable.Rows(0).Item(19).ToString : Else : Me.ComboBox34.Text = "" : End If
                If Not IsDBNull(MiDataTable.Rows(0).Item(22).ToString) Then : Me.ComboBox35.Text = MiDataTable.Rows(0).Item(22).ToString : Else : Me.ComboBox35.Text = "" : End If
                If Not IsDBNull(MiDataTable.Rows(0).Item(25).ToString) Then : Me.ComboBox36.Text = MiDataTable.Rows(0).Item(25).ToString : Else : Me.ComboBox36.Text = "" : End If
                Me.ComboBox3.Text = MiDataTable.Rows(0).Item(27).ToString           'TIPO
                Me.ComboBox4.Text = MiDataTable.Rows(0).Item(32).ToString           'CARGO
                Me.CheckBox1.Checked = MiDataTable.Rows(0).Item(43).ToString        'JEFE
                Me.ComboBox5.Text = MiDataTable.Rows(0).Item(34).ToString           'GP
                Me.ComboBox6.Text = MiDataTable.Rows(0).Item(36).ToString           'GR
                Me.ComboBox7.Text = MiDataTable.Rows(0).Item(38).ToString           'CATEGORIA SALARIAL
                If MiDataTable.Rows(0).Item(42).ToString = 1 Then                   'SITUACION
                    Me.ComboBox8.Text = "OCUPADO"
                End If
                If MiDataTable.Rows(0).Item(42).ToString = 2 Then
                    Me.ComboBox8.Text = "VACANTE"
                End If
                If MiDataTable.Rows(0).Item(42).ToString = 3 Then
                    Me.ComboBox8.Text = "CONGELADO"
                End If
                If MiDataTable.Rows(0).Item(42).ToString = 4 Then
                    Me.ComboBox8.Text = "."
                End If
                Me.CheckBox2.Checked = MiDataTable.Rows(0).Item(44).ToString        'MIXTA
                Me.CheckBox3.Checked = MiDataTable.Rows(0).Item(45).ToString        'MILITAR
                Me.CheckBox4.Checked = MiDataTable.Rows(0).Item(46).ToString        'PAME

                Me.ComboBox32.Text = MiDataTable.Rows(0).Item(48).ToString        'ESTABLECIMIENTO
                '------------------------------------------------------------------------------------
                Dim cadenaUBIC As String
                cadenaUBIC = Me.ComboBox1.Text
                If ComboBox2.Text = "" Then
                    GoTo SALIR
                Else
                    cadenaUBIC = cadenaUBIC & "\ " & Me.ComboBox2.Text
                End If
                If ComboBox29.Text = "" Then
                    GoTo SALIR
                Else
                    cadenaUBIC = cadenaUBIC & "\ " & Me.ComboBox29.Text
                End If
                If ComboBox30.Text = "" Then
                    GoTo SALIR
                Else
                    cadenaUBIC = cadenaUBIC & "\ " & Me.ComboBox30.Text
                End If
                If ComboBox31.Text = "" Then
                    GoTo SALIR
                Else
                    cadenaUBIC = cadenaUBIC & "\ " & Me.ComboBox31.Text
                End If
SALIR:
                Me.TextBox39.Text = cadenaUBIC                                      'UBICACION
                '------------------------------------------------------------------------------------
                Me.TextBox40.Text = MiDataTable.Rows(0).Item(32).ToString           'CARGO
                Me.TextBox41.Text = MiDataTable.Rows(0).Item(36).ToString           'GRADO REAL

                Me.TextBox35.Text = vID_M_P                                         'ID_M_P 
                Me.TextBox36.Text = vID_M_C                                         'ID_M_C
                Me.Cursor = Cursors.Default
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub BUSCAR_VISTA_MAESTRO_DE_PERSONAS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                Me.Cursor = Cursors.WaitCursor
                Me.DateTimePicker1.Value = MiDataTable.Rows(0).Item(1).ToString
                Me.DateTimePicker2.Value = MiDataTable.Rows(0).Item(2).ToString

                If Val(MiDataTable.Rows(0).Item(3).ToString) <> 0 Then
                    Me.DateTimePicker3.Value = MiDataTable.Rows(0).Item(3).ToString
                Else
                    Me.DateTimePicker3.Checked = False
                End If
                Me.TextBox5.Text = MiDataTable.Rows(0).Item(4).ToString
                Me.TextBox6.Text = MiDataTable.Rows(0).Item(5).ToString
                Me.TextBox30.Text = MiDataTable.Rows(0).Item(6).ToString
                Me.ComboBox9.Text = MiDataTable.Rows(0).Item(9).ToString
                Me.TextBox9.Text = MiDataTable.Rows(0).Item(12).ToString
                Me.TextBox42.Text = MiDataTable.Rows(0).Item(12).ToString
                Me.TextBox10.Text = MiDataTable.Rows(0).Item(13).ToString
                Me.TextBox11.Text = MiDataTable.Rows(0).Item(14).ToString
                Me.TextBox12.Text = MiDataTable.Rows(0).Item(15).ToString
                Me.TextBox13.Text = MiDataTable.Rows(0).Item(16).ToString
                Me.TextBox7.Text = MiDataTable.Rows(0).Item(10).ToString
                Me.TextBox8.Text = MiDataTable.Rows(0).Item(11).ToString
                Me.TextBox14.Text = MiDataTable.Rows(0).Item(17).ToString
                Me.TextBox15.Text = MiDataTable.Rows(0).Item(18).ToString
                Me.TextBox16.Text = MiDataTable.Rows(0).Item(19).ToString
                Me.TextBox17.Text = MiDataTable.Rows(0).Item(20).ToString
                Me.TextBox18.Text = MiDataTable.Rows(0).Item(21).ToString
                Me.TextBox21.Text = MiDataTable.Rows(0).Item(22).ToString
                Me.TextBox20.Text = MiDataTable.Rows(0).Item(23).ToString
                Me.TextBox19.Text = MiDataTable.Rows(0).Item(24).ToString
                Me.TextBox22.Text = MiDataTable.Rows(0).Item(25).ToString
                Me.TextBox23.Text = MiDataTable.Rows(0).Item(26).ToString
                Me.ComboBox10.Text = MiDataTable.Rows(0).Item(28).ToString
                Me.ComboBox11.Text = MiDataTable.Rows(0).Item(30).ToString
                Me.ComboBox13.Text = MiDataTable.Rows(0).Item(32).ToString
                Me.ComboBox12.Text = MiDataTable.Rows(0).Item(34).ToString
                Me.ComboBox15.Text = MiDataTable.Rows(0).Item(36).ToString
                Me.ComboBox14.Text = MiDataTable.Rows(0).Item(38).ToString
                Me.ComboBox16.Text = MiDataTable.Rows(0).Item(40).ToString
                Me.TextBox24.Text = MiDataTable.Rows(0).Item(41).ToString
                Me.ComboBox19.Text = MiDataTable.Rows(0).Item(43).ToString
                Me.ComboBox18.Text = MiDataTable.Rows(0).Item(45).ToString
                Me.ComboBox17.Text = MiDataTable.Rows(0).Item(47).ToString
                Me.TextBox25.Text = MiDataTable.Rows(0).Item(48).ToString
                Me.ComboBox20.Text = MiDataTable.Rows(0).Item(50).ToString
                Me.ComboBox21.Text = MiDataTable.Rows(0).Item(52).ToString
                'Me.TextBox26.Text = MiDataTable.Rows(0).Item(53).ToString
                Me.TextBox27.Text = MiDataTable.Rows(0).Item(54).ToString
                Me.ComboBox23.Text = MiDataTable.Rows(0).Item(56).ToString
                Me.ComboBox22.Text = MiDataTable.Rows(0).Item(58).ToString

                Me.CheckBox11.Checked = MiDataTable.Rows(0).Item(92).ToString()

                Me.TextBox31.Text = "" : Me.TextBox32.Text = "" : Me.TextBox33.Text = ""
                Me.TextBox31.Text = Format(Val(MiDataTable.Rows(0).Item(64).ToString), "##,##0.00")
                Me.TextBox32.Text = Format(Val(MiDataTable.Rows(0).Item(65).ToString), "##,##0.00")
                Me.TextBox33.Text = Format(VAL(MiDataTable.Rows(0).Item(66).ToString), "##,##0.00")
                'FOTO
                Call CARGAR_FOTO()
                '---
                Me.TextBox1.Text = MiDataTable.Rows(0).Item(7).ToString
                Me.TextBox2.Text = MiDataTable.Rows(0).Item(4).ToString
                Me.TextBox3.Text = Mid(MiDataTable.Rows(0).Item(2).ToString, 1, 10)
                '---
                Me.ComboBox24.Text = MiDataTable.Rows(0).Item(68).ToString
                Me.ComboBox25.Text = MiDataTable.Rows(0).Item(70).ToString
                Me.TextBox34.Text = MiDataTable.Rows(0).Item(71).ToString
                Me.ComboBox26.Text = MiDataTable.Rows(0).Item(73).ToString

                Me.CheckBox6.Checked = MiDataTable.Rows(0).Item(74).ToString
                Me.CheckBox7.Checked = MiDataTable.Rows(0).Item(75).ToString
                Me.CheckBox8.Checked = MiDataTable.Rows(0).Item(78).ToString
                Me.ComboBox27.Text = MiDataTable.Rows(0).Item(77).ToString

                If Not IsDBNull(MiDataTable.Rows(0).Item(80).ToString) Then : Me.TextBox44.Text = MiDataTable.Rows(0).Item(80).ToString : Else : Me.TextBox44.Text = "" : End If
                'CALCULO VACACIONAL
                If Not IsDBNull(MiDataTable.Rows(0).Item(81).ToString) Then : Me.CheckBox9.Checked = MiDataTable.Rows(0).Item(81).ToString : Else : Me.CheckBox9.Checked = False : End If
                'D_N_COMPETENCIA
                If Not IsDBNull(MiDataTable.Rows(0).Item(83).ToString) Then : Me.ComboBox28.Text = MiDataTable.Rows(0).Item(83).ToString : Else : Me.ComboBox28.Text = "" : End If
                'SERVICIOS_PROFESIONALES
                If Not IsDBNull(MiDataTable.Rows(0).Item(84).ToString) Then : Me.CheckBox10.Checked = MiDataTable.Rows(0).Item(84).ToString : Else : Me.CheckBox10.Checked = False : End If
                'PREFIJO
                If Not IsDBNull(MiDataTable.Rows(0).Item(85).ToString) Then : Me.TextBox28.Text = MiDataTable.Rows(0).Item(85).ToString : Else : Me.TextBox28.Text = "" : End If
                'ADMINISTRADOR DE RECURSOS
                Me.ComboBox33.Text = MiDataTable.Rows(0).Item(87).ToString

                Call BUSCAR_VISTA_MAESTRO_DE_CARGOS_2()
                Me.TextBox35.Text = vID_M_P
                Me.TextBox36.Text = vID_M_C
                Me.Cursor = Cursors.Default
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub BUSCAR_VISTA_MAESTRO_DE_CARGOS_2()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [VISTA MAESTRO DE CARGOS] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_C", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                Me.Cursor = Cursors.WaitCursor
                '------------------------------------------------------------------------------------
                Dim cadenaUBIC As String
                cadenaUBIC = Me.ComboBox1.Text
                If ComboBox2.Text = "" Then
                    GoTo SALIR
                Else
                    cadenaUBIC = cadenaUBIC & "\ " & Me.ComboBox2.Text
                End If
                If ComboBox29.Text = "" Then
                    GoTo SALIR
                Else
                    cadenaUBIC = cadenaUBIC & "\ " & Me.ComboBox29.Text
                End If
                If ComboBox30.Text = "" Then
                    GoTo SALIR
                Else
                    cadenaUBIC = cadenaUBIC & "\ " & Me.ComboBox30.Text
                End If
                If ComboBox31.Text = "" Then
                    GoTo SALIR
                Else
                    cadenaUBIC = cadenaUBIC & "\ " & Me.ComboBox31.Text
                End If
SALIR:
                Me.TextBox39.Text = cadenaUBIC                                      'UBICACION
                '------------------------------------------------------------------------------------
                Me.TextBox40.Text = MiDataTable.Rows(0).Item(23).ToString           'CARGO
                Me.TextBox41.Text = MiDataTable.Rows(0).Item(27).ToString           'GRADO REAL
                vID_M_C = MiDataTable.Rows(0).Item(0).ToString
                Me.Cursor = Cursors.Default
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_FOTO()
        Dim DT As New DataTable
        Dim DA As New SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE PERSONAS] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P"
            DA = New SqlDataAdapter(CADENAsql, CxRRHH)
            DA.Fill(DT)
            If DT.Rows.Count > 0 Then
                Dim ms As New System.IO.MemoryStream()
                Dim imageBuffer() As Byte = CType(DT.Rows(0).Item("FOTO"), Byte())
                ms = New System.IO.MemoryStream(imageBuffer)
                Me.PictureBox4.Image = Nothing
                Me.PictureBox4.Image = (Image.FromStream(ms))
                Me.PictureBox4.BackgroundImageLayout = ImageLayout.Stretch
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
        DT.Clear()
        DT.Reset()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call VERIFICAR_ACCESOS("075") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Expedientes_Buscar_Activos.ShowDialog()
        If CERRAR = False Then
            Dim tp As TabPage = TabGeneral.TabPages(0)
            'TabGeneral.SelectedTab = tp
            'Call SELECCIONAR_CADENA()
            '------------------------------------------------------------------------------------
            tp = TabGeneral.TabPages(1)
            TabGeneral.SelectedTab = tp
            Me.Label4.Text = Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Text
            Call NO_ACTIVAR()
            Call SELECCIONAR_CADENA()
            '------------------------------------------------------------------------------------
            Call BUSCAR_UBICACION()
            '------------------------------------------------------------------------------------
            Call DESACTIVAR_TODO()
            TabGeneral.Focus()
        End If
    End Sub
    Private Sub Button49_Click(sender As Object, e As EventArgs) Handles Button49.Click
        Call VERIFICAR_ACCESOS("098") : If HAY_ACCESO = False Then : Exit Sub : End If
        If Me.TextBox35.Text = "" Or Me.TextBox35.Text = "0" Then
            MsgBox("El registro que desea Editar debe estar visible en Pantalla", vbInformation, "Mensaje del Sistema")
            Me.Button1.Focus()
        Else
            Call ACTIVAR_01()
        End If
    End Sub
    Private Sub ACTIVAR_01()
        Me.TextBox4.ReadOnly = True : Me.TextBox37.ReadOnly = True : Me.TextBox38.ReadOnly = True
        Me.ComboBox1.Enabled = False : Me.ComboBox2.Enabled = False : Me.ComboBox3.Enabled = True
        Me.ComboBox4.Enabled = False : Me.CheckBox1.Enabled = True : Me.ComboBox5.Enabled = True
        Me.ComboBox6.Enabled = True : Me.ComboBox7.Enabled = True : Me.ComboBox8.Enabled = False : Me.ComboBox32.Enabled = True
        Me.CheckBox2.Enabled = False : Me.CheckBox3.Enabled = False : Me.CheckBox4.Enabled = False
        Me.Button49.Enabled = False
        Me.Button50.Enabled = True
        Me.Button48.Enabled = True
        '-------------------------------------
        'BLOQUEAR CONTROLES TAB : 0
        BLOQUEO = 0
        Call BLOQUEAR_OBJETOS()
        '-------------------------------------
        Me.ComboBox5.Focus()
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Call VERIFICAR_ACCESOS("100") : If HAY_ACCESO = False Then : Exit Sub : End If
        If Me.TextBox35.Text = "" Or Me.TextBox35.Text = "0" Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Me.Button1.Focus()
        Else
            Call ACTIVAR_02()
        End If
    End Sub
    Private Sub ACTIVAR_03()
        Me.TextBox26.ReadOnly = False
        Me.TextBox29.ReadOnly = False
        Me.TextBox43.ReadOnly = False
        Me.DateTimePicker8.Enabled = True
        'Me.DateTimePicker8.Checked = False
        Me.DateTimePicker9.Enabled = True
        'Me.DateTimePicker9.Checked = False
        Me.Button45.Enabled = False  'EDITAR
        Me.Button46.Enabled = True   'GUARDAR
        Me.Button3.Enabled = True   'CANCELAR
        '-------------------------------------
        'BLOQUEAR CONTROLES TAB : 2
        BLOQUEO = 2
        Call BLOQUEAR_OBJETOS()
        '-------------------------------------
        Me.TextBox26.Focus()
    End Sub
    Private Sub ACTIVAR_02()
        Me.DateTimePicker1.Enabled = True : Me.DateTimePicker2.Enabled = True : Me.DateTimePicker3.Enabled = True
        Me.TextBox5.ReadOnly = False : Me.TextBox6.ReadOnly = False : Me.TextBox30.ReadOnly = False
        Me.TextBox9.ReadOnly = False : Me.TextBox10.ReadOnly = False : Me.TextBox12.ReadOnly = False
        Me.TextBox13.ReadOnly = False : Me.TextBox11.ReadOnly = False : Me.TextBox7.ReadOnly = False
        Me.TextBox8.ReadOnly = False : Me.TextBox14.ReadOnly = False : Me.TextBox15.ReadOnly = False
        Me.TextBox16.ReadOnly = False : Me.TextBox17.ReadOnly = False : Me.TextBox18.ReadOnly = False
        Me.TextBox21.ReadOnly = False : Me.TextBox20.ReadOnly = False : Me.TextBox19.ReadOnly = False
        Me.TextBox22.ReadOnly = False : Me.TextBox23.ReadOnly = False : Me.TextBox24.ReadOnly = False
        Me.TextBox25.ReadOnly = False : Me.TextBox26.ReadOnly = False : Me.TextBox27.ReadOnly = False
        Me.TextBox34.ReadOnly = False : Me.CheckBox5.Enabled = True : Me.CheckBox11.Enabled = True
        Me.TextBox44.ReadOnly = False : Me.Button47.Enabled = True
        Me.TextBox28.ReadOnly = False : Me.Button63.Enabled = True
        Me.CheckBox6.Enabled = True
        Me.CheckBox7.Enabled = True
        Me.CheckBox9.Enabled = True
        Me.CheckBox10.Enabled = True
        Me.CheckBox8.Enabled = True
        Me.Button2.Enabled = True   'EXAMINAR
        Me.Button5.Enabled = False  'EDITAR
        Me.Button4.Enabled = True   'GUARDAR
        Me.Button8.Enabled = True   'CANCELAR
        '-------------------------------------
        'BLOQUEAR CONTROLES TAB : 1
        BLOQUEO = 1
        Call BLOQUEAR_OBJETOS()
        '-------------------------------------
        Me.DateTimePicker1.Focus()
    End Sub
    Private Sub DESACTIVAR_TODO()
        '0
        Me.TextBox4.ReadOnly = True
        Me.TextBox37.ReadOnly = True
        Me.TextBox38.ReadOnly = True
        Me.ComboBox1.Enabled = False
        Me.ComboBox2.Enabled = False
        Me.ComboBox3.Enabled = False
        Me.ComboBox4.Enabled = False
        Me.CheckBox1.Enabled = False
        Me.ComboBox5.Enabled = False
        Me.ComboBox6.Enabled = False
        Me.ComboBox7.Enabled = False
        Me.ComboBox8.Enabled = False
        Me.CheckBox2.Enabled = False
        Me.CheckBox3.Enabled = False
        Me.CheckBox4.Enabled = False
        Me.Button49.Enabled = True
        Me.Button50.Enabled = False
        Me.Button48.Enabled = False
        '1
        Me.DateTimePicker1.Enabled = False : Me.DateTimePicker2.Enabled = False : Me.DateTimePicker3.Enabled = False
        Me.TextBox5.ReadOnly = True : Me.TextBox6.ReadOnly = True : Me.TextBox30.ReadOnly = True
        Me.TextBox9.ReadOnly = True : Me.TextBox10.ReadOnly = True : Me.TextBox12.ReadOnly = True
        Me.TextBox13.ReadOnly = True : Me.TextBox11.ReadOnly = True : Me.TextBox7.ReadOnly = True
        Me.TextBox8.ReadOnly = True : Me.TextBox14.ReadOnly = True : Me.TextBox15.ReadOnly = True
        Me.TextBox16.ReadOnly = True : Me.TextBox17.ReadOnly = True : Me.TextBox18.ReadOnly = True
        Me.TextBox21.ReadOnly = True : Me.TextBox20.ReadOnly = True : Me.TextBox19.ReadOnly = True
        Me.TextBox22.ReadOnly = True : Me.TextBox23.ReadOnly = True : Me.TextBox24.ReadOnly = True
        Me.TextBox25.ReadOnly = True : Me.TextBox26.ReadOnly = True : Me.TextBox27.ReadOnly = True
        Me.TextBox34.ReadOnly = True : Me.CheckBox5.Enabled = False : Me.CheckBox11.Enabled = False
        Me.TextBox44.ReadOnly = True : Me.Button47.Enabled = False
        Me.TextBox28.ReadOnly = True : Me.Button63.Enabled = False
        Me.CheckBox6.Enabled = False
        Me.CheckBox7.Enabled = False
        Me.CheckBox9.Enabled = False
        Me.CheckBox10.Enabled = False
        Me.CheckBox8.Enabled = False
        Me.Button2.Enabled = False  'EXAMINAR
        Me.Button5.Enabled = True   'EDITAR
        Me.Button4.Enabled = False  'GUARDAR
        Me.Button8.Enabled = False  'CANCELAR
        '2
        Me.TextBox26.ReadOnly = True : Me.TextBox29.ReadOnly = True : Me.TextBox43.ReadOnly = True
        Me.DateTimePicker8.Enabled = False : Me.DateTimePicker9.Enabled = False
        Me.Button45.Enabled = True   'EDITAR
        Me.Button46.Enabled = False  'GUARDAR
        Me.Button3.Enabled = False  'CANCELAR
        '-------------------------------------
        Call DESBLOQUEAR_OBJETOS()
        '-------------------------------------
    End Sub
    Private Sub DESBLOQUEAR_OBJETOS()
        Me.TabGeneral.TabPages(0).Enabled = True
        Me.TabGeneral.TabPages(1).Enabled = True
        Me.TabGeneral.TabPages(2).Enabled = True
        Me.TabGeneral.TabPages(3).Enabled = True
        Me.TabGeneral.TabPages(4).Enabled = True
        Me.TabGeneral.TabPages(5).Enabled = True
        Me.TabGeneral.TabPages(6).Enabled = True
        Me.TabGeneral.TabPages(7).Enabled = True
        Me.TabGeneral.TabPages(8).Enabled = True
        Me.TabGeneral.TabPages(9).Enabled = True
        Me.TabGeneral.TabPages(10).Enabled = True
        Me.TabGeneral.TabPages(11).Enabled = True
        Me.TabGeneral.TabPages(12).Enabled = True
        Me.TabGeneral.TabPages(13).Enabled = True
        Me.TabGeneral.TabPages(14).Enabled = True
        If Val(Me.TextBox35.Text) <> 0 Then
            Me.Button11.Enabled = True : Me.Button10.Enabled = True : Me.Button9.Enabled = True
            Me.Button14.Enabled = True : Me.Button13.Enabled = True : Me.Button12.Enabled = True
            Me.Button17.Enabled = True : Me.Button16.Enabled = True : Me.Button15.Enabled = True
            Me.Button20.Enabled = True : Me.Button19.Enabled = True : Me.Button18.Enabled = True
            Me.Button23.Enabled = True : Me.Button22.Enabled = True : Me.Button21.Enabled = True
            Me.Button26.Enabled = True : Me.Button25.Enabled = True : Me.Button24.Enabled = True
            Me.Button29.Enabled = True : Me.Button28.Enabled = True : Me.Button27.Enabled = True
            Me.Button32.Enabled = True : Me.Button31.Enabled = True : Me.Button30.Enabled = True
            Me.Button35.Enabled = True : Me.Button34.Enabled = True : Me.Button33.Enabled = True
            Me.Button38.Enabled = True : Me.Button37.Enabled = True : Me.Button36.Enabled = True : Me.Button3.Enabled = True
            Me.Button41.Enabled = True : Me.Button40.Enabled = True : Me.Button39.Enabled = True
            Me.Button44.Enabled = True : Me.Button51.Enabled = True : Me.Button43.Enabled = True : Me.Button42.Enabled = True
            Me.Button57.Enabled = True : Me.Button56.Enabled = True : Me.Button55.Enabled = True : Me.Button58.Enabled = True
            Me.Button54.Enabled = True : Me.Button53.Enabled = True : Me.Button52.Enabled = True
            Me.Button59.Enabled = True : Me.Button60.Enabled = True : Me.Button61.Enabled = True
            Me.Button62.Enabled = True
            Me.DateTimePicker4.Enabled = True
            Me.DateTimePicker5.Enabled = True
        Else
            Me.Button11.Enabled = False : Me.Button10.Enabled = False : Me.Button9.Enabled = False
            Me.Button14.Enabled = False : Me.Button13.Enabled = False : Me.Button12.Enabled = False
            Me.Button17.Enabled = False : Me.Button16.Enabled = False : Me.Button15.Enabled = False
            Me.Button20.Enabled = False : Me.Button19.Enabled = False : Me.Button18.Enabled = False
            Me.Button23.Enabled = False : Me.Button22.Enabled = False : Me.Button21.Enabled = False
            Me.Button26.Enabled = False : Me.Button25.Enabled = False : Me.Button24.Enabled = False
            Me.Button29.Enabled = False : Me.Button28.Enabled = False : Me.Button27.Enabled = False
            Me.Button32.Enabled = False : Me.Button31.Enabled = False : Me.Button30.Enabled = False
            Me.Button35.Enabled = False : Me.Button34.Enabled = False : Me.Button33.Enabled = False
            Me.Button38.Enabled = False : Me.Button37.Enabled = False : Me.Button36.Enabled = False : Me.Button3.Enabled = False
            Me.Button41.Enabled = False : Me.Button40.Enabled = False : Me.Button39.Enabled = False
            Me.Button44.Enabled = False : Me.Button51.Enabled = False : Me.Button43.Enabled = False : Me.Button42.Enabled = False
            Me.Button57.Enabled = False : Me.Button56.Enabled = False : Me.Button55.Enabled = False : Me.Button58.Enabled = False
            Me.Button54.Enabled = False : Me.Button53.Enabled = False : Me.Button52.Enabled = False
            Me.Button59.Enabled = False : Me.Button60.Enabled = False : Me.Button61.Enabled = False
            Me.Button62.Enabled = False
            Me.DateTimePicker4.Enabled = False
            Me.DateTimePicker5.Enabled = False
        End If
    End Sub
    Dim BLOQUEO As Integer
    Private Sub BLOQUEAR_OBJETOS()
        If BLOQUEO = 0 Then : Me.TabGeneral.TabPages(0).Enabled = True : Else : Me.TabGeneral.TabPages(0).Enabled = False : End If
        If BLOQUEO = 1 Then : Me.TabGeneral.TabPages(1).Enabled = True : Else : Me.TabGeneral.TabPages(1).Enabled = False : End If
        If BLOQUEO = 2 Then : Me.TabGeneral.TabPages(2).Enabled = True : Else : Me.TabGeneral.TabPages(2).Enabled = False : End If
        'If BLOQUEO = 3 Then : Me.TabGeneral.TabPages(3).Enabled = True : Else : Me.TabGeneral.TabPages(3).Enabled = False : End If
        'If BLOQUEO = 4 Then : Me.TabGeneral.TabPages(4).Enabled = True : Else : Me.TabGeneral.TabPages(4).Enabled = False : End If
        'If BLOQUEO = 5 Then : Me.TabGeneral.TabPages(5).Enabled = True : Else : Me.TabGeneral.TabPages(5).Enabled = False : End If
        'If BLOQUEO = 6 Then : Me.TabGeneral.TabPages(6).Enabled = True : Else : Me.TabGeneral.TabPages(6).Enabled = False : End If
        'If BLOQUEO = 7 Then : Me.TabGeneral.TabPages(7).Enabled = True : Else : Me.TabGeneral.TabPages(7).Enabled = False : End If
        'If BLOQUEO = 8 Then : Me.TabGeneral.TabPages(8).Enabled = True : Else : Me.TabGeneral.TabPages(8).Enabled = False : End If
        'If BLOQUEO = 9 Then : Me.TabGeneral.TabPages(9).Enabled = True : Else : Me.TabGeneral.TabPages(9).Enabled = False : End If
        'If BLOQUEO = 10 Then : Me.TabGeneral.TabPages(10).Enabled = True : Else : Me.TabGeneral.TabPages(10).Enabled = False : End If
        'If BLOQUEO = 11 Then : Me.TabGeneral.TabPages(11).Enabled = True : Else : Me.TabGeneral.TabPages(11).Enabled = False : End If
        'If BLOQUEO = 12 Then : Me.TabGeneral.TabPages(12).Enabled = True : Else : Me.TabGeneral.TabPages(12).Enabled = False : End If
        'If BLOQUEO = 13 Then : Me.TabGeneral.TabPages(13).Enabled = True : Else : Me.TabGeneral.TabPages(13).Enabled = False : End If
        'If BLOQUEO = 14 Then : Me.TabGeneral.TabPages(14).Enabled = True : Else : Me.TabGeneral.TabPages(14).Enabled = False : End If
        'If BLOQUEO = 15 Then : Me.TabGeneral.TabPages(15).Enabled = True : Else : Me.TabGeneral.TabPages(15).Enabled = False : End If
        'If BLOQUEO = 16 Then : Me.TabGeneral.TabPages(16).Enabled = True : Else : Me.TabGeneral.TabPages(16).Enabled = False : End If
        'If BLOQUEO = 17 Then : Me.TabGeneral.TabPages(17).Enabled = True : Else : Me.TabGeneral.TabPages(17).Enabled = False : End If
        'If BLOQUEO = 18 Then : Me.TabGeneral.TabPages(18).Enabled = True : Else : Me.TabGeneral.TabPages(18).Enabled = False : End If
        'If BLOQUEO = 19 Then : Me.TabGeneral.TabPages(19).Enabled = True : Else : Me.TabGeneral.TabPages(19).Enabled = False : End If
        'If BLOQUEO = 20 Then : Me.TabGeneral.TabPages(20).Enabled = True : Else : Me.TabGeneral.TabPages(20).Enabled = False : End If
        'If BLOQUEO = 21 Then : Me.TabGeneral.TabPages(21).Enabled = True : Else : Me.TabGeneral.TabPages(21).Enabled = False : End If
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim tp As TabPage = TabGeneral.TabPages(1)
        TabGeneral.SelectedTab = tp
        Me.Label4.Text = Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Text
        Call NO_ACTIVAR()
        Call SELECCIONAR_CADENA()
        TabGeneral.Focus()
        Call DESACTIVAR_TODO()
    End Sub
    Private Sub Button48_Click(sender As Object, e As EventArgs) Handles Button48.Click
        Dim tp As TabPage = TabGeneral.TabPages(0)
        TabGeneral.SelectedTab = tp
        Me.Label4.Text = Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Text
        Call NO_ACTIVAR()
        Call SELECCIONAR_CADENA()
        TabGeneral.Focus()
        Call DESACTIVAR_TODO()
    End Sub
    Private Sub DateTimePicker1_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub DateTimePicker2_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub DateTimePicker3_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox6.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox30_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox30.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox9.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox9.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox10.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox12_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox12.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox13_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox13.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox11_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox11.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox24_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox24.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox25_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox25.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox7.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox14_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox14.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox17_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox17.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox16_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox16.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox15_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox15.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox18_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox18.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox21_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox21.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub TextBox19_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox19.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub TextBox20_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox20.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox12_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox12.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox23_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox23.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox10.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox11_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox11.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox13_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox13.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox22_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox22.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox15_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox15.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox14_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox14.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox16_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox16.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox24_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox24.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox25_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox25.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox18_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox18.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox17_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox17.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox19_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox19.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox20_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox20.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox21_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox21.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox27_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox27.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub CheckBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox22_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox22.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox23_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox23.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox34_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox34.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    'Private Sub TextBox28_KeyDown(sender As Object, e As KeyEventArgs)
    '    If e.KeyCode = Keys.Enter Then
    '        e.Handled = True
    '        SendKeys.Send("{TAB}")
    '    End If
    'End Sub
    'Private Sub TextBox29_KeyDown(sender As Object, e As KeyEventArgs)
    '    If e.KeyCode = Keys.Enter Then
    '        e.Handled = True
    '        SendKeys.Send("{TAB}")
    '    End If
    'End Sub
    Private Sub TextBox33_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox33.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox32_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox32.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub CheckBox11_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox31_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox31.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Dim CONTROL_FOTO As Boolean
    Dim data() As Byte 'arreglo de bytes el cual contedra la imagen
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call VERIFICAR_ACCESOS("102") : If HAY_ACCESO = False Then : Exit Sub : End If
        Dim ios As FileStream 'Manejo de archivos
        Try
            Me.OpenFileDialog1.Filter = "Imagenes (JPG)|*.jpg" 'filtro de archivos del OpenFileDialog 
            If Me.OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.Cancel Then ' en caso de que se aplaste el boton cancelar salga y no haga nada
                CONTROL_FOTO = False
                Exit Sub
            Else
                ios = New FileStream(Me.OpenFileDialog1.FileName, FileMode.Open, FileAccess.Read) 'instanciamos en Stream indicandole la ruta del arvhivo,el modo de acceso y si es de lectura o escritura
                ReDim data(ios.Length) 'llenamos el arreglo
                ios.Read(data, 0, CInt(ios.Length)) 'llenamos el arreglo
                Me.PictureBox4.SizeMode = PictureBoxSizeMode.StretchImage 'establecemos como se visualiza la imagen
                Me.PictureBox4.Load(Me.OpenFileDialog1.FileName) 'visualizamos abriendo el archivo seleccionado
                CONTROL_FOTO = True
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Mensaje del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification) 'en caso de error muestre un mensaje
        End Try
    End Sub
    Dim EXISTE_CODIGO As Boolean
    Dim CEDULA1, EMPLEADO1 As String
    Private Sub VERIFICAR_SI_EXISTE_CODIGO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [MAESTRO DE PERSONAS] WHERE CODIGO = '" & Me.TextBox5.Text & "' AND ID_M_P <> " & Me.TextBox35.Text & " AND ID_ESTADO_P = 1  ORDER BY CODIGO", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                EMPLEADO1 = MiDataTable.Rows(0).Item(5).ToString & " " & MiDataTable.Rows(0).Item(6).ToString
                CEDULA1 = MiDataTable.Rows(0).Item(11).ToString
                EXISTE_CODIGO = True
            Else
                EMPLEADO1 = ""
                CEDULA1 = ""
                EXISTE_CODIGO = False
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim vcMIXTA, vcMILITAR, vcPAME As Boolean
    Private Sub BUSCAR_ESTRUCTURA_CARGO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [MAESTRO DE CARGOS] WHERE ID_M_C = '" & Me.TextBox36.Text & "'", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                vcMIXTA = MiDataTable.Rows(0).Item(15).ToString
                vcMILITAR = MiDataTable.Rows(0).Item(16).ToString
                vcPAME = MiDataTable.Rows(0).Item(17).ToString
            Else
                vcMIXTA = False
                vcMILITAR = False
                vcPAME = False
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button50_Click(sender As Object, e As EventArgs) Handles Button50.Click 'GUARDAR
        Call VERIFICAR_ACCESOS("099") : If HAY_ACCESO = False Then : Exit Sub : End If
        If Me.TextBox35.Text = "" Or Me.TextBox35.Text = "0" Or Me.TextBox36.Text = "" Or Me.TextBox36.Text = "0" Then
            MsgBox("No se encuentra el registro", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Esta seguro de Actualizar el registro en Pantalla?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            If Len(Me.ComboBox5.SelectedValue) = 0 Then
                MsgBox("Debe seleccionar el Grado Plantilla Correctamente", vbInformation, "Mensaje del Sistema")
                Me.ComboBox5.Focus()
                Exit Sub
            End If
            If Len(Me.ComboBox6.SelectedValue) = 0 Then
                MsgBox("Debe seleccionar el Grado Real Correctamente", vbInformation, "Mensaje del Sistema")
                Me.ComboBox6.Focus()
                Exit Sub
            End If
            If Len(Me.ComboBox7.SelectedValue) = 0 Then
                MsgBox("Debe seleccionar la Categoria Salarial Correctamente", vbInformation, "Mensaje del Sistema")
                Me.ComboBox7.Focus()
                Exit Sub
            End If
            If Len(Me.ComboBox32.SelectedValue) = 0 Then
                MsgBox("Debe seleccionar el Establecimiento Correctamente", vbInformation, "Mensaje del Sistema")
                Me.ComboBox32.Focus()
                Exit Sub
            End If
            If Len(Me.ComboBox3.SelectedValue) = 0 Then
                MsgBox("Debe seleccionar el Tipo de Estructura de Cargo Correctamente", vbInformation, "Mensaje del Sistema")
                Me.ComboBox3.Focus()
                Exit Sub
            End If
            Call ACTUALIZAR_REGISTRO_ESTRUCTURA()
            Dim tp As TabPage = TabGeneral.TabPages(0)
            TabGeneral.SelectedTab = tp
            Me.Label4.Text = Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Text
            Call NO_ACTIVAR()
            Call SELECCIONAR_CADENA()
            TabGeneral.Focus()
            Call DESACTIVAR_TODO()
        End If
    End Sub
    Private Sub ACTUALIZAR_REGISTRO_ESTRUCTURA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim SQL As String
        Try
            SQL = "UPDATE [MAESTRO DE CARGOS] SET ID_T_ES = " & Me.ComboBox3.SelectedValue & ", ID_GP = " & Me.ComboBox5.SelectedValue & ", ID_GR = " & Me.ComboBox6.SelectedValue & ", ID_CAT_SALARIAL = " & Me.ComboBox7.SelectedValue & ", JEFE = '" & Me.CheckBox1.Checked & "', ID_ESTABLECIMIENTO = " & Me.ComboBox32.SelectedValue & " WHERE ID_M_C = " & Me.TextBox36.Text & ""
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
    Private Sub Button46_Click_1(sender As Object, e As EventArgs) Handles Button46.Click
        Call VERIFICAR_ACCESOS("104") : If HAY_ACCESO = False Then : Exit Sub : End If
        MENSAJE = MsgBox("¿Esta seguro de Actualizar el registro en Pantalla?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            If Me.TextBox26.Text = "" Then
                Me.TextBox26.Text = ""
            End If
            If Me.TextBox29.Text = "" Then
                MsgBox("Debe digitar los Nombres y Apellidos del Beneficiario", vbInformation, "Mensaje del Sistema")
                Me.TextBox29.SelectAll()
                Me.TextBox29.Focus()
            End If
            If Me.TextBox43.Text = "" Then
                MsgBox("Debe digitar el Pais de Emision de la Cédula", vbInformation, "Mensaje del Sistema")
                Me.TextBox43.SelectAll()
                Me.TextBox43.Focus()
            End If
            If Me.DateTimePicker8.Checked = False Then
                MsgBox("Debe digitar una Fecha de Emisión de Cédula válida", vbInformation, "Mensaje del Sistema")
                Me.DateTimePicker8.Focus()
            End If
            If Me.DateTimePicker9.Checked = False Then
                MsgBox("Debe digitar una Fecha de Vencimiento de Cédula válida", vbInformation, "Mensaje del Sistema")
                Me.DateTimePicker9.Focus()
            End If
            Call ACTUALIZAR_REGISTRO_BANCO()
            Dim tp As TabPage = TabGeneral.TabPages(2)
            TabGeneral.SelectedTab = tp
            Me.Label4.Text = Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Text
            Call NO_ACTIVAR()
            Call SELECCIONAR_CADENA()
            TabGeneral.Focus()
            Call DESACTIVAR_TODO()
        End If
    End Sub
    Private Sub ACTUALIZAR_REGISTRO_BANCO()
        Dim gCUENTABANCO As Object = If(Me.TextBox26.Text <> "", "Convert(NVARCHAR(20), '" + Trim(Me.TextBox26.Text) + "', 23)", "NULL")
        Dim fFEC_EMI_ID As String
        If Me.DateTimePicker8.Checked = True Then
            fFEC_EMI_ID = Mid(Me.DateTimePicker8.Value, 1, 10)
        Else
            fFEC_EMI_ID = "NULL"
        End If
        Dim fFEC_VEN_ID As String
        If Me.DateTimePicker9.Checked = True Then
            fFEC_VEN_ID = Mid(Me.DateTimePicker9.Value, 1, 10)
        Else
            fFEC_VEN_ID = "NULL"
        End If

        Dim fechaemisionid As Object = If(fFEC_EMI_ID <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFEC_EMI_ID + "', 105), 23)", "NULL")
        Dim fechavencimientoid As Object = If(fFEC_VEN_ID <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFEC_VEN_ID + "', 105), 23)", "NULL")

        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim SQL As String
        Try

            SQL = "UPDATE [MAESTRO DE PERSONAS] SET N_C_BANCO = " & gCUENTABANCO & ", NA_BENEFICIARIO = '" & Me.TextBox29.Text & "', PAIS_EMISION_ID = '" & Me.TextBox43.Text & "', FEC_EMISION_ID= " & fechaemisionid & ", FEC_VENCIMIENTO_ID = " & fechavencimientoid & " WHERE ID_M_P = " & Me.TextBox35.Text & ""
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
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click   'GUARDAR
        Call VERIFICAR_ACCESOS("101") : If HAY_ACCESO = False Then : Exit Sub : End If
        MENSAJE = MsgBox("¿Esta seguro de Actualizar el registro en Pantalla?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            If Me.TextBox5.Text = "" Then
                MsgBox("El código de Empleado no es válido", vbInformation, "Mensaje del Sistema")
                Me.TextBox5.SelectAll()
                Me.TextBox5.Focus()
                Exit Sub
            End If
            TextBox5.Text = Format(Val(Me.TextBox5.Text), "0000000000")
            Call VERIFICAR_SI_EXISTE_CODIGO()
            If EXISTE_CODIGO = True Then
                Me.TextBox5.SelectAll()
                Me.TextBox5.Focus()
                MENSAJE = MsgBox("El Código de Empleado que intenta ingresar ya existe, se encuentra asignado a " & CEDULA1 & " " & EMPLEADO1 & ", ¿Desea Generar un Código Nuevo?", vbInformation + vbYesNo, "Mensaje del Sistema")
                If MENSAJE = vbYes Then
                    Call GENERAR_CODIGO_EMPLEADO()
                    Me.TextBox6.Focus()
                Else
                    Exit Sub
                End If
            End If
            If Me.TextBox6.Text = "" Then
                MsgBox("Los Apellidos no son válidos", vbInformation, "Mensaje del Sistema")
                Me.TextBox6.SelectAll()
                Me.TextBox6.Focus()
                Exit Sub
            End If
            If Me.TextBox30.Text = "" Then
                MsgBox("Los Nombres no son válidos", vbInformation, "Mensaje del Sistema")
                Me.TextBox30.SelectAll()
                Me.TextBox30.Focus()
                Exit Sub
            End If
            If Len(Me.ComboBox9.SelectedValue) = 0 Then
                MsgBox("Debe seleccionar el Tipo de Sexo Correctamente", vbInformation, "Mensaje del Sistema")
                Me.ComboBox9.Focus()
                Exit Sub
            End If
            If Me.TextBox9.Text = "" Then
                Me.TextBox9.Text = ""
            End If
            If Me.TextBox10.Text = "" Then
                Me.TextBox10.Text = ""
            End If
            If Me.TextBox12.Text = "" Then
                Me.TextBox12.Text = ""
            End If
            If Me.TextBox13.Text = "" Then
                Me.TextBox13.Text = ""
            End If
            If Me.TextBox11.Text = "" Then
                Me.TextBox11.Text = ""
            End If
            If Len(Me.ComboBox24.SelectedValue) = 0 Then
                MsgBox("Debe seleccionar la Categoria de Licencia Correctamente", vbInformation, "Mensaje del Sistema")
                Me.ComboBox24.Focus()
                Exit Sub
            End If
            If Len(Me.ComboBox25.SelectedValue) = 0 Then
                MsgBox("Debe seleccionar el Tipo Sanguineo Correctamente", vbInformation, "Mensaje del Sistema")
                Me.ComboBox25.Focus()
                Exit Sub
            End If
            If Me.TextBox7.Text = "" Then
                Me.TextBox7.Text = ""
            End If
            If Me.TextBox8.Text = "" Then
                Me.TextBox8.Text = ""
            End If
            If Me.TextBox14.Text = "" Then
                Me.TextBox14.Text = ""
            End If
            If Me.TextBox15.Text = "" Then
                Me.TextBox15.Text = ""
            End If
            If Me.TextBox16.Text = "" Then
                Me.TextBox16.Text = ""
            End If
            If Me.TextBox17.Text = "" Then
                Me.TextBox17.Text = ""
            End If
            If Me.TextBox18.Text = "" Then
                Me.TextBox18.Text = ""
            End If
            If Me.TextBox21.Text = "" Then
                Me.TextBox21.Text = ""
            End If
            If Me.TextBox20.Text = "" Then
                Me.TextBox20.Text = ""
            End If
            If Me.TextBox19.Text = "" Then
                Me.TextBox19.Text = ""
            End If
            If Me.TextBox22.Text = "" Then
                Me.TextBox22.Text = ""
            End If
            If Me.TextBox23.Text = "" Then
                Me.TextBox23.Text = ""
            End If
            If Len(Me.ComboBox10.SelectedValue) = 0 Then
                MsgBox("Debe seleccionar el Tipo de Tez Correctamente", vbInformation, "Mensaje del Sistema")
                Me.ComboBox10.Focus()
                Exit Sub
            End If
            If Len(Me.ComboBox11.SelectedValue) = 0 Then
                MsgBox("Debe seleccionar el Tipo de Cabello Correctamente", vbInformation, "Mensaje del Sistema")
                Me.ComboBox11.Focus()
                Exit Sub
            End If
            If Len(Me.ComboBox11.SelectedValue) = 0 Then
                MsgBox("Debe seleccionar el Tipo de Cabello Correctamente", vbInformation, "Mensaje del Sistema")
                Me.ComboBox11.Focus()
                Exit Sub
            End If
            If Len(Me.ComboBox12.SelectedValue) = 0 Then
                MsgBox("Debe seleccionar el Nivel Profesional Correctamente", vbInformation, "Mensaje del Sistema")
                Me.ComboBox12.Focus()
                Exit Sub
            End If
            If Len(Me.ComboBox15.SelectedValue) = 0 Then
                MsgBox("Debe seleccionar la Nacionalidad de Nacimiento Correctamente", vbInformation, "Mensaje del Sistema")
                Me.ComboBox15.Focus()
                Exit Sub
            End If
            If Len(Me.ComboBox14.SelectedValue) = 0 Then
                MsgBox("Debe seleccionar el Departamento de Nacimiento Correctamente", vbInformation, "Mensaje del Sistema")
                Me.ComboBox14.Focus()
                Exit Sub
            End If
            If Len(Me.ComboBox16.SelectedValue) = 0 Then
                MsgBox("Debe seleccionar el Municipio de Nacimiento Correctamente", vbInformation, "Mensaje del Sistema")
                Me.ComboBox16.Focus()
                Exit Sub
            End If
            If Me.TextBox24.Text = "" Then
                Me.TextBox24.Text = ""
            End If
            If Len(Me.ComboBox19.SelectedValue) = 0 Then
                MsgBox("Debe seleccionar la Nacionalidad Domiciliar Correctamente", vbInformation, "Mensaje del Sistema")
                Me.ComboBox19.Focus()
                Exit Sub
            End If
            If Len(Me.ComboBox18.SelectedValue) = 0 Then
                MsgBox("Debe seleccionar el Departamento Domiciliar Correctamente", vbInformation, "Mensaje del Sistema")
                Me.ComboBox18.Focus()
                Exit Sub
            End If
            If Len(Me.ComboBox17.SelectedValue) = 0 Then
                MsgBox("Debe seleccionar el Municipio Domiciliar Correctamente", vbInformation, "Mensaje del Sistema")
                Me.ComboBox17.Focus()
                Exit Sub
            End If
            If Me.TextBox25.Text = "" Then
                Me.TextBox25.Text = ""
            End If
            If Len(Me.ComboBox20.SelectedValue) = 0 Then
                MsgBox("Debe seleccionar el Estado Civil Correctamente", vbInformation, "Mensaje del Sistema")
                Me.ComboBox20.Focus()
                Exit Sub
            End If
            If Len(Me.ComboBox21.SelectedValue) = 0 Then
                MsgBox("Debe seleccionar el Tipo de Horario Correctamente", vbInformation, "Mensaje del Sistema")
                Me.ComboBox21.Focus()
                Exit Sub
            End If
            If Me.TextBox26.Text = "" Then
                Me.TextBox26.Text = ""
            End If
            If Me.TextBox27.Text = "" Then
                Me.TextBox27.Text = ""
            End If
            If Len(Me.ComboBox23.SelectedValue) = 0 Then
                MsgBox("Debe seleccionar el Tipo de Afiliacion Correctamente", vbInformation, "Mensaje del Sistema")
                Me.ComboBox23.Focus()
                Exit Sub
            End If
            If Len(Me.ComboBox22.SelectedValue) = 0 Then
                MsgBox("Debe seleccionar el Tipo de Plaza Correctamente", vbInformation, "Mensaje del Sistema")
                Me.ComboBox22.Focus()
                Exit Sub
            End If
            If Me.TextBox34.Text = "" Then
                Me.TextBox34.Text = ""
            End If
            If Len(Me.ComboBox26.SelectedValue) = 0 Then
                MsgBox("Debe seleccionar el Ipss", vbInformation, "Mensaje del Sistema")
                Me.ComboBox26.Focus()
                Exit Sub
            End If
            If Len(Me.ComboBox27.SelectedValue) = 0 Then
                MsgBox("Debe seleccionar el la Funcionalidad Correctamente", vbInformation, "Mensaje del Sistema")
                Me.ComboBox17.Focus()
                Exit Sub
            End If

            If Me.TextBox44.Text = "" Then
                MsgBox("Debe Digitar Nombres y Apellidos del Jefe Inmediato", vbInformation, "Mensaje del Sistema")
                Me.TextBox44.Focus()
                Exit Sub
            End If
            If Me.TextBox28.Text = "" Then
                MsgBox("Debe Digitar el Prefijo a Utilizar", vbInformation, "Mensaje del Sistema")
                Me.TextBox28.Focus()
                Exit Sub
            End If
            If Len(Me.ComboBox33.SelectedValue) = 0 Then
                MsgBox("Debe seleccionar el Administrador del Recurso Correctamente", vbInformation, "Mensaje del Sistema")
                Me.ComboBox33.Focus()
                Exit Sub
            End If
            Dim AÑOS As Integer
            Call Clases.BUSCAR_FECHA_Y_HORA()
            AÑOS = (FECHA_SERVIDOR.Year) - (Me.DateTimePicker1.Value.Year)
            If AÑOS < 16 Then
                MsgBox("Error, la edad del Empleado es de " & AÑOS & " años, por favor verifique la Fecha de Nacimiento", vbInformation, "Mensaje del Sistema")
                Me.DateTimePicker1.Focus()
                Exit Sub
            End If
            Call ACTUALIZAR_REGISTRO()
            Dim tp As TabPage = TabGeneral.TabPages(1)
            TabGeneral.SelectedTab = tp
            Me.Label4.Text = Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Text
            Call NO_ACTIVAR()
            Call SELECCIONAR_CADENA()
            TabGeneral.Focus()
            Call DESACTIVAR_TODO()
        End If
    End Sub
    Private Sub GENERAR_CODIGO_EMPLEADO()
        Dim CODIGO_EMPLEADO As Integer
        CODIGO_EMPLEADO = 0
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            MiDataAdapter = New SqlDataAdapter("SELECT CODIGO FROM [MAESTRO DE PERSONAS] ORDER BY CODIGO", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                For I = 0 To MiDataTable.Rows.Count - 1
                    CODIGO_EMPLEADO = CInt(MiDataTable.Rows(I).Item(0).ToString)
                Next
                CODIGO_EMPLEADO += 1
            Else
                CODIGO_EMPLEADO = 1
            End If
            Me.TextBox5.Text = CODIGO_EMPLEADO
            Me.TextBox5.Text = Format(CInt(Me.TextBox5.Text), "0000000000")
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ACTUALIZAR_REGISTRO()
        CONTROL_FOTO = True
        Dim gDIRECCION_1 As Object = If(Me.TextBox7.Text <> "", "Convert(NVARCHAR(MAX), '" + Trim(Me.TextBox7.Text) + "', 23)", "NULL")
        Dim gDIRECCION_2 As Object = If(Me.TextBox8.Text <> "", "Convert(NVARCHAR(MAX), '" + Trim(Me.TextBox8.Text) + "', 23)", "NULL")
        Dim gNSS As Object = If(Me.TextBox9.Text <> "", "Convert(NVARCHAR(50), '" + Trim(Me.TextBox9.Text) + "', 23)", "NULL")
        Dim gN_CEDULA As Object = If(Me.TextBox10.Text <> "", "Convert(NVARCHAR(50), '" + Trim(Me.TextBox10.Text) + "', 23)", "NULL")
        Dim gN_LIC As Object = If(Me.TextBox11.Text <> "", "Convert(NVARCHAR(50), '" + Trim(Me.TextBox11.Text) + "', 23)", "NULL")
        Dim gN_MINSA As Object = If(Me.TextBox12.Text <> "", "Convert(NVARCHAR(50), '" + Trim(Me.TextBox12.Text) + "', 23)", "NULL")
        Dim gLIC_IMAGEN As Object = If(Me.TextBox13.Text <> "", "Convert(NVARCHAR(50), '" + Trim(Me.TextBox13.Text) + "', 23)", "NULL")
        Dim gTELEFONO_1 As Object = If(Me.TextBox14.Text <> "", "Convert(NVARCHAR(50), '" + Trim(Me.TextBox14.Text) + "', 23)", "NULL")
        Dim gTELEFONO_2 As Object = If(Me.TextBox15.Text <> "", "Convert(NVARCHAR(50), '" + Trim(Me.TextBox15.Text) + "', 23)", "NULL")
        Dim gCONVENCIONAL As Object = If(Me.TextBox16.Text <> "", "Convert(NVARCHAR(50), '" + Trim(Me.TextBox16.Text) + "', 23)", "NULL")
        Dim gCORREO_1 As Object = If(Me.TextBox17.Text <> "", "Convert(NVARCHAR(100), '" + Trim(Me.TextBox17.Text) + "', 23)", "NULL")
        Dim gCORREO_2 As Object = If(Me.TextBox18.Text <> "", "Convert(NVARCHAR(100), '" + Trim(Me.TextBox18.Text) + "', 23)", "NULL")
        Dim gWHATSAPP As Object = If(Me.TextBox21.Text <> "", "Convert(NVARCHAR(50), '" + Trim(Me.TextBox21.Text) + "', 23)", "NULL")
        Dim gFACEBOOK As Object = If(Me.TextBox20.Text <> "", "Convert(NVARCHAR(50), '" + Trim(Me.TextBox20.Text) + "', 23)", "NULL")
        Dim gTWITTER As Object = If(Me.TextBox19.Text <> "", "Convert(NVARCHAR(50), '" + Trim(Me.TextBox19.Text) + "', 23)", "NULL")
        Dim gPESO As Object = If(Me.TextBox22.Text <> "", "Convert(FLOAT, '" + Trim(Me.TextBox22.Text) + "', 23)", "NULL")
        Dim gESTATURA As Object = If(Me.TextBox23.Text <> "", "Convert(FLOAT, '" + Trim(Me.TextBox23.Text) + "', 23)", "NULL")
        Dim gLUGAR_N As Object = If(Me.TextBox24.Text <> "", "Convert(NVARCHAR(200), '" + Trim(Me.TextBox24.Text) + "', 23)", "NULL")
        Dim gLUGAR_D As Object = If(Me.TextBox25.Text <> "", "Convert(NVARCHAR(200), '" + Trim(Me.TextBox25.Text) + "', 23)", "NULL")
        Dim gN_C_BANCO As Object = If(Me.TextBox26.Text <> "", "Convert(NVARCHAR(200), '" + Trim(Me.TextBox26.Text) + "', 23)", "NULL")
        Dim gEMPLEADO_ANT As Object = If(Me.TextBox27.Text <> "", "Convert(NVARCHAR(200), '" + Trim(Me.TextBox27.Text) + "', 23)", "NULL")
        Dim gOBSERVACIONES As Object = If(Me.TextBox34.Text <> "", "Convert(NVARCHAR(MAX), '" + Trim(Me.TextBox34.Text) + "', 23)", "NULL")
        Dim gJEFE_REAL As Object = If(Me.TextBox44.Text <> "", "Convert(NVARCHAR(200), '" + Trim(Me.TextBox44.Text) + "', 23)", "NULL")
        Dim gPREFIJO As Object = If(Me.TextBox28.Text <> "", "Convert(NVARCHAR(200), '" + Trim(Me.TextBox28.Text) + "', 23)", "NULL")

        Dim fFEC_ING_EN As String
        Dim fFEC_NAC As String = Mid(Me.DateTimePicker1.Value, 1, 10)
        Dim fFEC_ING As String = Mid(Me.DateTimePicker2.Value, 1, 10)
        If Me.DateTimePicker3.Checked = True Then
            fFEC_ING_EN = Mid(Me.DateTimePicker3.Value, 1, 10)
        Else
            fFEC_ING_EN = "NULL"
        End If
        Dim fechaNacimiento = "Convert(VARCHAR(10), Convert(Date, '" + fFEC_NAC + "', 105), 23)"
        Dim fechaingreso = "Convert(VARCHAR(10), Convert(Date, '" + fFEC_ING + "', 105), 23)"
        Dim fechaingresoEN As Object = If(fFEC_ING_EN <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFEC_ING_EN + "', 105), 23)", "NULL")

        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim SQL As String
        Dim ms As New System.IO.MemoryStream() 'creando el memoryStream
        Try
            If CONTROL_FOTO = True Then
                Me.PictureBox4.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
            Else
                Me.PictureBox7.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
            End If
            SQL = "UPDATE [MAESTRO DE PERSONAS] SET FEC_NAC = " & fechaNacimiento & ", FEC_ING_PAME = " & fechaingreso & ", FEC_ING_EN = " & fechaingresoEN & ", CODIGO = '" & Me.TextBox5.Text & "', APELLIDOS = '" & Me.TextBox6.Text & "', NOMBRES = '" & Me.TextBox30.Text & "', ID_T_SEXO = " & Me.ComboBox9.SelectedValue & ", DIRECCION_1 = " & gDIRECCION_1 & ", DIRECCION_2 = " & gDIRECCION_2 & ", NSS = " & gNSS & ", N_CEDULA = " & gN_CEDULA & ", N_LIC = " & gN_LIC & ", N_MINSA = " & gN_MINSA & ", LIC_IMAGEN = " & gLIC_IMAGEN & ", TELEFONO_1 = " & gTELEFONO_1 & ", TELEFONO_2 = " & gTELEFONO_2 & ", CONVENCIONAL = " & gCONVENCIONAL & ", CORREO_1 = " & gCORREO_1 & ", CORREO_2 = " & gCORREO_2 & ", WHATSAPP = " & gWHATSAPP & ", FACEBOOK = " & gFACEBOOK & ", TWITTER = " & gTWITTER & ", PESO = " & gPESO & ", ESTATURA = " & gESTATURA & ", ID_TEZ = " & Me.ComboBox10.SelectedValue & ", ID_CABELLO = " & Me.ComboBox11.SelectedValue & ", ID_N_ACADEMICO = " & Me.ComboBox13.SelectedValue & ", ID_N_PROFESIONAL = " & Me.ComboBox12.SelectedValue & ", ID_NACIONALIDAD_N = " & Me.ComboBox15.SelectedValue & ", ID_DPTO_N = " & Me.ComboBox14.SelectedValue & ", ID_M_N = " & Me.ComboBox16.SelectedValue & ", LUGAR_N = " & gLUGAR_N & ", ID_NACIONALIDAD_D = " & Me.ComboBox19.SelectedValue & ", ID_DPTO_D = " & Me.ComboBox18.SelectedValue & ", ID_M_D = " & Me.ComboBox17.SelectedValue & ", LUGAR_D = " & gLUGAR_D & ", ID_E_CIVIL = " & Me.ComboBox20.SelectedValue & ", ID_T_HORARIO = " & Me.ComboBox21.SelectedValue & ", N_C_BANCO = " & gN_C_BANCO & ", EMPLEADO_ANT = " & gEMPLEADO_ANT & ", ID_T_AFILIACION = " & Me.ComboBox23.SelectedValue & ", ID_T_PLAZA = " & Me.ComboBox22.SelectedValue & ", D_N1 = NULL, D_N2 = NULL, FOTO = @data, BECADO = '" & Me.CheckBox5.Checked & "', ID_C_LIC = " & Me.ComboBox24.SelectedValue & ", ID_T_SANGUINEO = " & Me.ComboBox25.SelectedValue & ", OBSERVACIONES = " & gOBSERVACIONES & ", ID_IPSS = " & Me.ComboBox26.SelectedValue & ", ESPECIALISTA = '" & Me.CheckBox6.Checked & "', SUBESPECIALISTA = '" & Me.CheckBox7.Checked & "', ID_F = " & Me.ComboBox27.SelectedValue & ", EN_SITUACION = '" & Me.CheckBox8.Checked & "', JEFE_REAL = " & gJEFE_REAL & ", CALCULO_VACACIONAL = '" & Me.CheckBox9.Checked & "', ID_N_COMPETENCIA = " & Me.ComboBox28.SelectedValue & ", SERVICIOS_PROFESIONALES = '" & Me.CheckBox10.Checked & "', PREFIJO = " & gPREFIJO & ", ID_ADMIN = " & Me.ComboBox33.SelectedValue & ", PRACTICANTE = '" & Me.CheckBox11.Checked & "' WHERE ID_M_P = " & Me.TextBox35.Text & ""
            Dim cmd As New SqlCommand(SQL, CxRRHH)
            cmd.Parameters.Add(New SqlParameter("@data", ms.GetBuffer()))
            cmd.ExecuteNonQuery()
            cmd = Nothing
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
        CONTROL_FOTO = False
    End Sub
    'VISTA MAESTRO DE NUCLEO FAMILIAR
    Private Sub DGV1_Click(sender As Object, e As EventArgs) Handles DGV1.Click
        If Me.DGV1.RowCount > 0 Then
            For I = 0 To Me.DGV1.RowCount - 1
                If DGV1.Rows(I).Selected = True Then
                    If Not IsDBNull(Me.DGV1.Item(0, I).Value) Then : vmnfID_NF = Me.DGV1.Item(0, I).Value : Else : vmnfID_NF = 0 : End If
                    If Not IsDBNull(Me.DGV1.Item(1, I).Value) Then : vmnfID_M_P = Me.DGV1.Item(1, I).Value : Else : vmnfID_M_P = 0 : End If
                    If Not IsDBNull(Me.DGV1.Item(2, I).Value) Then : vmnfID_PARENTELA = Me.DGV1.Item(2, I).Value : Else : vmnfID_PARENTELA = 0 : End If
                    If Not IsDBNull(Me.DGV1.Item(3, I).Value) Then : vmnfD_PARENTELA = Me.DGV1.Item(3, I).Value : Else : vmnfD_PARENTELA = "" : End If
                    If Not IsDBNull(Me.DGV1.Item(4, I).Value) Then : vmnfAPELLIDOS_NOMBRES = Me.DGV1.Item(4, I).Value : Else : vmnfAPELLIDOS_NOMBRES = "" : End If
                    If Not IsDBNull(Me.DGV1.Item(5, I).Value) Then : vmnfID_T_SEXO = Me.DGV1.Item(5, I).Value : Else : vmnfID_T_SEXO = 0 : End If
                    If Not IsDBNull(Me.DGV1.Item(6, I).Value) Then : vmnfD_T_SEXO = Me.DGV1.Item(6, I).Value : Else : vmnfD_T_SEXO = "" : End If
                    If Not IsDBNull(Me.DGV1.Item(7, I).Value) Then : vmnfFEC_NAC = Me.DGV1.Item(7, I).Value : Else : vmnfFEC_NAC = "" : End If
                    If Not IsDBNull(Me.DGV1.Item(8, I).Value) Then : vmnfEDAD = Val(Me.DGV1.Item(8, I).Value) : Else : vmnfEDAD = "" : End If
                    If Not IsDBNull(Me.DGV1.Item(9, I).Value) Then : vmnfID_N_ACADEMICO = Me.DGV1.Item(9, I).Value : Else : vmnfID_N_ACADEMICO = 0 : End If
                    If Not IsDBNull(Me.DGV1.Item(10, I).Value) Then : vmnfD_N_ACADEMICO = Me.DGV1.Item(10, I).Value : Else : vmnfD_N_ACADEMICO = "" : End If
                    If Not IsDBNull(Me.DGV1.Item(11, I).Value) Then : vmnfDIRECCION_D = Me.DGV1.Item(11, I).Value : Else : vmnfDIRECCION_D = "" : End If
                    If Not IsDBNull(Me.DGV1.Item(12, I).Value) Then : vmnfTELEFONO_1 = Me.DGV1.Item(12, I).Value : Else : vmnfTELEFONO_1 = "" : End If
                    If Not IsDBNull(Me.DGV1.Item(13, I).Value) Then : vmnfTELEFONO_2 = Me.DGV1.Item(13, I).Value : Else : vmnfTELEFONO_2 = "" : End If
                    If Not IsDBNull(Me.DGV1.Item(14, I).Value) Then : vmnfLUGAR_T = Me.DGV1.Item(14, I).Value : Else : vmnfLUGAR_T = "" : End If
                    If Not IsDBNull(Me.DGV1.Item(15, I).Value) Then : vmnfDIRECCION_T = Me.DGV1.Item(15, I).Value : Else : vmnfDIRECCION_T = "" : End If
                    If Not IsDBNull(Me.DGV1.Item(16, I).Value) Then : vmnfCARGO_T = Me.DGV1.Item(16, I).Value : Else : vmnfCARGO_T = "" : End If
                    If Not IsDBNull(Me.DGV1.Item(17, I).Value) Then : vmnfFALLECIDO = Me.DGV1.Item(17, I).Value : Else : vmnfFALLECIDO = "" : End If
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click   'ELIMINAR
        Call VERIFICAR_ACCESOS("107") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmnfID_NF = 0 Then
            MsgBox("Debe seleccionar el registro que desea Eliminar", vbInformation, "Mensaje del Sistema")
            Me.DGV1.Focus()
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Esta seguro de Eliminar el registro seleccionado?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Conexion.ABD_RRHH()
            Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
            Try
                MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE NUCLEO FAMILIAR] WHERE ID_NF = " & CInt(vmnfID_NF) & ""
                MiSqlCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
            Call CARGAR_DATAGRIDVIEW_1()
            Call ARMAR_DATAGRIDVIEW_1()
        Else
            vmnfID_NF = 0
        End If
    End Sub
    Private Sub DGV1_DoubleClick(sender As Object, e As EventArgs) Handles DGV1.DoubleClick 'EDITAR
        Call VERIFICAR_ACCESOS("106") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmnfID_NF = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Editar_Nucleo_Familiar.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE NUCLEO FAMILIAR] WHERE ID_M_P = " & Me.TextBox35.Text & " ORDER BY ID_M_P, ID_NF"
            Call CARGAR_DATAGRIDVIEW_1()
            Call ARMAR_DATAGRIDVIEW_1()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_1()
        End If
    End Sub
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click 'EDITAR
        Call VERIFICAR_ACCESOS("106") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmnfID_NF = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Editar_Nucleo_Familiar.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE NUCLEO FAMILIAR] WHERE ID_M_P = " & Me.TextBox35.Text & " ORDER BY ID_M_P, ID_NF"
            Call CARGAR_DATAGRIDVIEW_1()
            Call ARMAR_DATAGRIDVIEW_1()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_1()
        End If
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_1()
        Dim I As Integer
        For I = 0 To Me.DGV1.RowCount - 1
            If Me.DGV1.Item(0, I).Value = vmnfID_NF Then
                Me.DGV1.Rows(I).Selected = True
                Me.DGV1.CurrentCell = Me.DGV1.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click 'AGREGAR
        Call VERIFICAR_ACCESOS("105") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Agregar_Nucleo_Familiar.ShowDialog()
    End Sub
    'MAESTRO DE EXPERIENCIA LABORAL
    Private Sub DGV2_Click(sender As Object, e As EventArgs) Handles DGV2.Click
        If Me.DGV2.RowCount > 0 Then
            For I = 0 To Me.DGV2.RowCount - 1
                If DGV2.Rows(I).Selected = True Then
                    If Not IsDBNull(Me.DGV2.Item(0, I).Value) Then : vmelID_EL = Me.DGV2.Item(0, I).Value : Else : vmelID_EL = 0 : End If
                    If Not IsDBNull(Me.DGV2.Item(1, I).Value) Then : vmelID_M_P = Me.DGV2.Item(1, I).Value : Else : vmelID_M_P = 0 : End If
                    If Not IsDBNull(Me.DGV2.Item(2, I).Value) Then : vmelEMPRESA_NOMBRE = Me.DGV2.Item(2, I).Value : Else : vmelEMPRESA_NOMBRE = "" : End If
                    If Not IsDBNull(Me.DGV2.Item(3, I).Value) Then : vmelEMPRESA_DIRECCION = Me.DGV2.Item(3, I).Value : Else : vmelEMPRESA_DIRECCION = "" : End If
                    If Not IsDBNull(Me.DGV2.Item(4, I).Value) Then : vmelEMPRESA_TELEFONO = Me.DGV2.Item(4, I).Value : Else : vmelEMPRESA_TELEFONO = "" : End If
                    If Not IsDBNull(Me.DGV2.Item(5, I).Value) Then : vmelEMPRESA_ACTIVIDAD = Me.DGV2.Item(5, I).Value : Else : vmelEMPRESA_ACTIVIDAD = "" : End If
                    If Not IsDBNull(Me.DGV2.Item(6, I).Value) Then : vmelCARGO_DESEM = Me.DGV2.Item(6, I).Value : Else : vmelCARGO_DESEM = "" : End If
                    If Not IsDBNull(Me.DGV2.Item(7, I).Value) Then : vmelJEFE_INMEDIATO = Me.DGV2.Item(7, I).Value : Else : vmelJEFE_INMEDIATO = "" : End If
                    If Not IsDBNull(Me.DGV2.Item(8, I).Value) Then : vmelSALARIO = Me.DGV2.Item(8, I).Value : Else : vmelSALARIO = "" : End If
                    If Not IsDBNull(Me.DGV2.Item(9, I).Value) Then : vmelFEC_INI = Me.DGV2.Item(9, I).Value : Else : vmelFEC_INI = "" : End If
                    If Not IsDBNull(Me.DGV2.Item(10, I).Value) Then : vmelFEC_FIN = Me.DGV2.Item(10, I).Value : Else : vmelFEC_FIN = "" : End If
                    If Not IsDBNull(Me.DGV2.Item(11, I).Value) Then : vmelCAUSA_RETIRO = Me.DGV2.Item(11, I).Value : Else : vmelCAUSA_RETIRO = "" : End If
                    If Not IsDBNull(Me.DGV2.Item(12, I).Value) Then : vmelOBLI_RESP = Me.DGV2.Item(12, I).Value : Else : vmelOBLI_RESP = "" : End If
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click   'ELIMINAR
        Call VERIFICAR_ACCESOS("110") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmelID_EL = 0 Then
            MsgBox("Debe seleccionar el registro que desea Eliminar", vbInformation, "Mensaje del Sistema")
            Me.DGV2.Focus()
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Esta seguro de Eliminar el registro seleccionado?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Conexion.ABD_RRHH()
            Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
            Try
                MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE EXPERIENCIA LABORAL] WHERE ID_EL = " & CInt(vmelID_EL) & ""
                MiSqlCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
            Call CARGAR_DATAGRIDVIEW_2()
            Call ARMAR_DATAGRIDVIEW_2()
        Else
            vmelID_EL = 0
        End If
    End Sub
    Private Sub DGV2_DoubleClick(sender As Object, e As EventArgs) Handles DGV2.DoubleClick 'EDITAR
        Call VERIFICAR_ACCESOS("109") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmelID_EL = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Editar_Experiencia_Laboral.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [MAESTRO DE EXPERIENCIA LABORAL] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_EL"
            Call CARGAR_DATAGRIDVIEW_2()
            Call ARMAR_DATAGRIDVIEW_2()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_2()
        End If
    End Sub
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click 'EDITAR
        Call VERIFICAR_ACCESOS("109") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmelID_EL = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Editar_Experiencia_Laboral.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [MAESTRO DE EXPERIENCIA LABORAL] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_EL"
            Call CARGAR_DATAGRIDVIEW_2()
            Call ARMAR_DATAGRIDVIEW_2()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_2()
        End If
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_2()
        Dim I As Integer
        For I = 0 To Me.DGV2.RowCount - 1
            If Me.DGV2.Item(0, I).Value = vmelID_EL Then
                Me.DGV2.Rows(I).Selected = True
                Me.DGV2.CurrentCell = Me.DGV2.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click 'AGREGAR
        Call VERIFICAR_ACCESOS("108") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Agregar_Experiencia_Laboral.ShowDialog()
    End Sub
    'MAESTRO DE REFERENCIAS PERSONALES
    Private Sub DGV3_Click(sender As Object, e As EventArgs) Handles DGV3.Click
        If Me.DGV3.RowCount > 0 Then
            For I = 0 To Me.DGV3.RowCount - 1
                If DGV3.Rows(I).Selected = True Then
                    If Not IsDBNull(Me.DGV3.Item(0, I).Value) Then : vmrID_RP = Me.DGV3.Item(0, I).Value : Else : vmrID_RP = 0 : End If
                    If Not IsDBNull(Me.DGV3.Item(1, I).Value) Then : vmrID_M_P = Me.DGV3.Item(1, I).Value : Else : vmrID_M_P = 0 : End If
                    If Not IsDBNull(Me.DGV3.Item(2, I).Value) Then : vmrAPELLIDOS_NOMBRES = Me.DGV3.Item(2, I).Value : Else : vmrAPELLIDOS_NOMBRES = "" : End If
                    If Not IsDBNull(Me.DGV3.Item(3, I).Value) Then : vmrDIRECCION = Me.DGV3.Item(3, I).Value : Else : vmrDIRECCION = "" : End If
                    If Not IsDBNull(Me.DGV3.Item(4, I).Value) Then : vmrOCUPACION = Me.DGV3.Item(4, I).Value : Else : vmrOCUPACION = "" : End If
                    If Not IsDBNull(Me.DGV3.Item(5, I).Value) Then : vmrTELEFONO = Me.DGV3.Item(5, I).Value : Else : vmrTELEFONO = "" : End If
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click   'ELIMINAR
        Call VERIFICAR_ACCESOS("113") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmrID_RP = 0 Then
            MsgBox("Debe seleccionar el registro que desea Eliminar", vbInformation, "Mensaje del Sistema")
            Me.DGV3.Focus()
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Esta seguro de Eliminar el registro seleccionado?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Conexion.ABD_RRHH()
            Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
            Try
                MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE REFERENCIAS PERSONALES] WHERE ID_RP = " & CInt(vmrID_RP) & ""
                MiSqlCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
            Call CARGAR_DATAGRIDVIEW_3()
            Call ARMAR_DATAGRIDVIEW_3()
        Else
            vmrID_RP = 0
        End If
    End Sub
    Private Sub DGV3_DoubleClick(sender As Object, e As EventArgs) Handles DGV3.DoubleClick 'EDITAR
        Call VERIFICAR_ACCESOS("112") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmrID_RP = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Editar_Referencias_Personales.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [MAESTRO DE REFERENCIAS PERSONALES] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_RP"
            Call CARGAR_DATAGRIDVIEW_3()
            Call ARMAR_DATAGRIDVIEW_3()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_3()
        End If
    End Sub
    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click 'EDITAR
        Call VERIFICAR_ACCESOS("112") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmrID_RP = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Editar_Referencias_Personales.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [MAESTRO DE REFERENCIAS PERSONALES] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_RP"
            Call CARGAR_DATAGRIDVIEW_3()
            Call ARMAR_DATAGRIDVIEW_3()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_3()
        End If
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_3()
        Dim I As Integer
        For I = 0 To Me.DGV3.RowCount - 1
            If Me.DGV3.Item(0, I).Value = vmrID_RP Then
                Me.DGV3.Rows(I).Selected = True
                Me.DGV3.CurrentCell = Me.DGV3.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click 'AGREGAR
        Call VERIFICAR_ACCESOS("111") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Agregar_Referencias_Personales.ShowDialog()
    End Sub
    'VISTA MAESTRO DE ESTUDIOS
    Private Sub DGV4_Click(sender As Object, e As EventArgs) Handles DGV4.Click
        If Me.DGV4.RowCount > 0 Then
            For I = 0 To Me.DGV4.RowCount - 1
                If DGV4.Rows(I).Selected = True Then
                    If Not IsDBNull(Me.DGV4.Item(0, I).Value) Then : vmesID_M_EST = Me.DGV4.Item(0, I).Value : Else : vmesID_M_EST = 0 : End If
                    If Not IsDBNull(Me.DGV4.Item(1, I).Value) Then : vmesID_M_P = Me.DGV4.Item(1, I).Value : Else : vmesID_M_P = 0 : End If
                    If Not IsDBNull(Me.DGV4.Item(2, I).Value) Then : vmesID_T_ESTUDIO = Me.DGV4.Item(2, I).Value : Else : vmesID_T_ESTUDIO = 0 : End If
                    If Not IsDBNull(Me.DGV4.Item(3, I).Value) Then : vmesD_T_ESTUDIO = Me.DGV4.Item(3, I).Value : Else : vmesD_T_ESTUDIO = "" : End If
                    If Not IsDBNull(Me.DGV4.Item(4, I).Value) Then : vmesID_E_MED = Me.DGV4.Item(4, I).Value : Else : vmesID_E_MED = 0 : End If
                    If Not IsDBNull(Me.DGV4.Item(5, I).Value) Then : vmesD_E_MED = Me.DGV4.Item(5, I).Value : Else : vmesD_E_MED = "" : End If
                    If Not IsDBNull(Me.DGV4.Item(6, I).Value) Then : vmesULT_GRADO_A = Me.DGV4.Item(6, I).Value : Else : vmesULT_GRADO_A = "" : End If
                    If Not IsDBNull(Me.DGV4.Item(7, I).Value) Then : vmesFECHA_ESTUDIO = Me.DGV4.Item(7, I).Value : Else : vmesFECHA_ESTUDIO = "" : End If
                    If Not IsDBNull(Me.DGV4.Item(8, I).Value) Then : vmesNOMBRE_ESTUDIO = Me.DGV4.Item(8, I).Value : Else : vmesNOMBRE_ESTUDIO = "" : End If
                    If Not IsDBNull(Me.DGV4.Item(9, I).Value) Then : vmesNOMBRE_C_ESTUDIO = Me.DGV4.Item(9, I).Value : Else : vmesNOMBRE_C_ESTUDIO = "" : End If
                    If Not IsDBNull(Me.DGV4.Item(10, I).Value) Then : vmesDIRECC_C_ESTUDIO = Me.DGV4.Item(10, I).Value : Else : vmesDIRECC_C_ESTUDIO = "" : End If
                    If Not IsDBNull(Me.DGV4.Item(11, I).Value) Then : vmesESTUDIO_F_PAIS = Me.DGV4.Item(11, I).Value : Else : vmesESTUDIO_F_PAIS = False : End If
                    If Not IsDBNull(Me.DGV4.Item(12, I).Value) Then : vmesID_PAIS = Me.DGV4.Item(12, I).Value : Else : vmesID_PAIS = 0 : End If
                    If Not IsDBNull(Me.DGV4.Item(13, I).Value) Then : vmesD_PAIS = Me.DGV4.Item(13, I).Value : Else : vmesD_PAIS = "" : End If
                    If Not IsDBNull(Me.DGV4.Item(14, I).Value) Then : vmesTITULO_OBT = Me.DGV4.Item(14, I).Value : Else : vmesTITULO_OBT = "" : End If
                    If Not IsDBNull(Me.DGV4.Item(15, I).Value) Then : vmesOBSERVACIONES = Me.DGV4.Item(15, I).Value : Else : vmesOBSERVACIONES = "" : End If
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click   'ELIMINAR
        Call VERIFICAR_ACCESOS("116") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmesID_M_EST = 0 Then
            MsgBox("Debe seleccionar el registro que desea Eliminar", vbInformation, "Mensaje del Sistema")
            Me.DGV4.Focus()
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Esta seguro de Eliminar el registro seleccionado?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Conexion.ABD_RRHH()
            Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
            Try
                MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE ESTUDIOS] WHERE ID_M_EST = " & CInt(vmesID_M_EST) & ""
                MiSqlCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
            Call CARGAR_DATAGRIDVIEW_4()
            Call ARMAR_DATAGRIDVIEW_4()
        Else
            vmesID_M_EST = 0
        End If
    End Sub
    Private Sub DGV4_DoubleClick(sender As Object, e As EventArgs) Handles DGV4.DoubleClick 'EDITAR
        Call VERIFICAR_ACCESOS("115") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmesID_M_EST = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Editar_Estudios.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE ESTUDIOS] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_M_EST"
            Call CARGAR_DATAGRIDVIEW_4()
            Call ARMAR_DATAGRIDVIEW_4()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_4()
        End If
    End Sub
    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click 'EDITAR
        Call VERIFICAR_ACCESOS("115") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmesID_M_EST = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Editar_Estudios.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE ESTUDIOS] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_M_EST"
            Call CARGAR_DATAGRIDVIEW_4()
            Call ARMAR_DATAGRIDVIEW_4()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_4()
        End If
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_4()
        Dim I As Integer
        For I = 0 To Me.DGV4.RowCount - 1
            If Me.DGV4.Item(0, I).Value = vmesID_M_EST Then
                Me.DGV4.Rows(I).Selected = True
                Me.DGV4.CurrentCell = Me.DGV4.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click 'AGREGAR
        Call VERIFICAR_ACCESOS("114") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Agregar_Estudios.ShowDialog()
    End Sub
    'VISTA MAESTRO DE HABILIDADES ADMIN
    Private Sub DGV5_Click(sender As Object, e As EventArgs) Handles DGV5.Click
        If Me.DGV5.RowCount > 0 Then
            For I = 0 To Me.DGV5.RowCount - 1
                If DGV5.Rows(I).Selected = True Then
                    If Not IsDBNull(Me.DGV5.Item(0, I).Value) Then : vmhID_H_ADMIN = Me.DGV5.Item(0, I).Value : Else : vmhID_H_ADMIN = 0 : End If
                    If Not IsDBNull(Me.DGV5.Item(1, I).Value) Then : vmhID_M_P = Me.DGV5.Item(1, I).Value : Else : vmhID_M_P = 0 : End If
                    If Not IsDBNull(Me.DGV5.Item(2, I).Value) Then : vmhID_HABILIDAD = Me.DGV5.Item(2, I).Value : Else : vmhID_HABILIDAD = 0 : End If
                    If Not IsDBNull(Me.DGV5.Item(3, I).Value) Then : vmhD_HABILIDAD = Me.DGV5.Item(3, I).Value : Else : vmhD_HABILIDAD = "" : End If
                    If Not IsDBNull(Me.DGV5.Item(4, I).Value) Then : vmhPORCENTAJE = Me.DGV5.Item(4, I).Value : Else : vmhPORCENTAJE = "" : End If
                    If Not IsDBNull(Me.DGV5.Item(5, I).Value) Then : vmhOBSERVACIONES = Me.DGV5.Item(5, I).Value : Else : vmhOBSERVACIONES = "" : End If
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click   'ELIMINAR
        Call VERIFICAR_ACCESOS("119") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmhID_H_ADMIN = 0 Then
            MsgBox("Debe seleccionar el registro que desea Eliminar", vbInformation, "Mensaje del Sistema")
            Me.DGV5.Focus()
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Esta seguro de Eliminar el registro seleccionado?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Conexion.ABD_RRHH()
            Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
            Try
                MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE HABILIDADES ADMIN] WHERE ID_H_ADMIN = " & CInt(vmhID_H_ADMIN) & ""
                MiSqlCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
            Call CARGAR_DATAGRIDVIEW_5()
            Call ARMAR_DATAGRIDVIEW_5()
        Else
            vmhID_H_ADMIN = 0
        End If
    End Sub
    Private Sub DGV5_DoubleClick(sender As Object, e As EventArgs) Handles DGV5.DoubleClick
        Call VERIFICAR_ACCESOS("118") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmhID_H_ADMIN = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Editar_Habilidades_Admin.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HABILIDADES ADMIN] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_H_ADMIN"
            Call CARGAR_DATAGRIDVIEW_5()
            Call ARMAR_DATAGRIDVIEW_5()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_5()
        End If
    End Sub
    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        Call VERIFICAR_ACCESOS("118") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmhID_H_ADMIN = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Editar_Habilidades_Admin.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HABILIDADES ADMIN] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_H_ADMIN"
            Call CARGAR_DATAGRIDVIEW_5()
            Call ARMAR_DATAGRIDVIEW_5()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_5()
        End If
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_5()
        Dim I As Integer
        For I = 0 To Me.DGV5.RowCount - 1
            If Me.DGV5.Item(0, I).Value = vmhID_H_ADMIN Then
                Me.DGV5.Rows(I).Selected = True
                Me.DGV5.CurrentCell = Me.DGV5.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        Call VERIFICAR_ACCESOS("117") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Agregar_Habilidades_Admin.ShowDialog()
    End Sub
    'VISTA MAESTRO DE IDIOMAS
    Private Sub DGV6_Click(sender As Object, e As EventArgs) Handles DGV6.Click
        If Me.DGV6.RowCount > 0 Then
            For I = 0 To Me.DGV6.RowCount - 1
                If DGV6.Rows(I).Selected = True Then
                    If Not IsDBNull(Me.DGV6.Item(0, I).Value) Then : vmiID_M_I = Me.DGV6.Item(0, I).Value : Else : vmiID_M_I = 0 : End If
                    If Not IsDBNull(Me.DGV6.Item(1, I).Value) Then : vmiID_M_P = Me.DGV6.Item(1, I).Value : Else : vmiID_M_P = 0 : End If
                    If Not IsDBNull(Me.DGV6.Item(2, I).Value) Then : vmiID_IDIOMA = Me.DGV6.Item(2, I).Value : Else : vmiID_IDIOMA = 0 : End If
                    If Not IsDBNull(Me.DGV6.Item(3, I).Value) Then : vmiD_IDIOMA = Me.DGV6.Item(3, I).Value : Else : vmiD_IDIOMA = "" : End If
                    If Not IsDBNull(Me.DGV6.Item(4, I).Value) Then : vmiHABLA = Me.DGV6.Item(4, I).Value : Else : vmiHABLA = 0 : End If
                    If Not IsDBNull(Me.DGV6.Item(5, I).Value) Then : vmiP_HABLA = Me.DGV6.Item(5, I).Value : Else : vmiP_HABLA = "" : End If
                    If Not IsDBNull(Me.DGV6.Item(6, I).Value) Then : vmiLEE = Me.DGV6.Item(6, I).Value : Else : vmiLEE = 0 : End If
                    If Not IsDBNull(Me.DGV6.Item(7, I).Value) Then : vmiPORCENTAJE_L = Me.DGV6.Item(7, I).Value : Else : vmiPORCENTAJE_L = "" : End If
                    If Not IsDBNull(Me.DGV6.Item(8, I).Value) Then : vmiESCRIBE = Me.DGV6.Item(8, I).Value : Else : vmiESCRIBE = 0 : End If
                    If Not IsDBNull(Me.DGV6.Item(9, I).Value) Then : vmiPORCENTAJE_E = Me.DGV6.Item(9, I).Value : Else : vmiPORCENTAJE_E = "" : End If
                    If Not IsDBNull(Me.DGV6.Item(10, I).Value) Then : vmiOBSERVACIONES = Me.DGV6.Item(10, I).Value : Else : vmiOBSERVACIONES = "" : End If
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click   'ELIMINAR
        Call VERIFICAR_ACCESOS("122") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmiID_M_I = 0 Then
            MsgBox("Debe seleccionar el registro que desea Eliminar", vbInformation, "Mensaje del Sistema")
            Me.DGV6.Focus()
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Esta seguro de Eliminar el registro seleccionado?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Conexion.ABD_RRHH()
            Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
            Try
                MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE IDIOMAS] WHERE ID_M_I = " & CInt(vmiID_M_I) & ""
                MiSqlCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
            Call CARGAR_DATAGRIDVIEW_6()
            Call ARMAR_DATAGRIDVIEW_6()
        Else
            vmiID_M_I = 0
        End If
    End Sub
    Private Sub DGV6_DoubleClick(sender As Object, e As EventArgs) Handles DGV6.DoubleClick 'EDITAR
        Call VERIFICAR_ACCESOS("121") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmiID_M_I = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Editar_Idioma.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE IDIOMAS] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_M_I"
            Call CARGAR_DATAGRIDVIEW_6()
            Call ARMAR_DATAGRIDVIEW_6()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_6()
        End If
    End Sub
    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click 'EDITAR
        Call VERIFICAR_ACCESOS("121") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmiID_M_I = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Editar_Idioma.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE IDIOMAS] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_M_I"
            Call CARGAR_DATAGRIDVIEW_6()
            Call ARMAR_DATAGRIDVIEW_6()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_6()
        End If
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_6()
        Dim I As Integer
        For I = 0 To Me.DGV6.RowCount - 1
            If Me.DGV6.Item(0, I).Value = vmiID_M_I Then
                Me.DGV6.Rows(I).Selected = True
                Me.DGV6.CurrentCell = Me.DGV6.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click 'AGREGAR
        Call VERIFICAR_ACCESOS("120") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Agregar_Idioma.ShowDialog()
    End Sub
    'VISTA MAESTRO DE CURVA Y TALLA
    Private Sub DGV7_Click(sender As Object, e As EventArgs) Handles DGV7.Click
        If Me.DGV7.RowCount > 0 Then
            For I = 0 To Me.DGV7.RowCount - 1
                If DGV7.Rows(I).Selected = True Then
                    If Not IsDBNull(Me.DGV7.Item(0, I).Value) Then : vmcytID_CT = Me.DGV7.Item(0, I).Value : Else : vmcytID_CT = 0 : End If
                    If Not IsDBNull(Me.DGV7.Item(1, I).Value) Then : vmcytID_M_P = Me.DGV7.Item(1, I).Value : Else : vmcytID_M_P = 0 : End If
                    If Not IsDBNull(Me.DGV7.Item(2, I).Value) Then : vmcytID_T_CT = Me.DGV7.Item(2, I).Value : Else : vmcytID_T_CT = 0 : End If
                    If Not IsDBNull(Me.DGV7.Item(3, I).Value) Then : vmcytD_T_CT = Me.DGV7.Item(3, I).Value : Else : vmcytD_T_CT = "" : End If
                    If Not IsDBNull(Me.DGV7.Item(4, I).Value) Then : vmcytTALLA = Me.DGV7.Item(4, I).Value : Else : vmcytTALLA = "" : End If
                    If Not IsDBNull(Me.DGV7.Item(5, I).Value) Then : vmcytOBSERVACIONES = Me.DGV7.Item(5, I).Value : Else : vmcytOBSERVACIONES = "" : End If
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click   'ELIMINAR
        Call VERIFICAR_ACCESOS("125") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmcytID_CT = 0 Then
            MsgBox("Debe seleccionar el registro que desea Eliminar", vbInformation, "Mensaje del Sistema")
            Me.DGV7.Focus()
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Esta seguro de Eliminar el registro seleccionado?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Conexion.ABD_RRHH()
            Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
            Try
                MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE CURVA Y TALLA] WHERE ID_CT = " & CInt(vmcytID_CT) & ""
                MiSqlCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
            Call CARGAR_DATAGRIDVIEW_7()
            Call ARMAR_DATAGRIDVIEW_7()
        Else
            vmcytID_CT = 0
        End If
    End Sub
    Private Sub DGV7_DoubleClick(sender As Object, e As EventArgs) Handles DGV7.DoubleClick 'EDITAR
        Call VERIFICAR_ACCESOS("124") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmcytID_CT = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Editar_Curva_y_Talla.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE CURVA Y TALLA] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_CT"
            Call CARGAR_DATAGRIDVIEW_7()
            Call ARMAR_DATAGRIDVIEW_7()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_7()
        End If
    End Sub
    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click 'EDITAR
        Call VERIFICAR_ACCESOS("124") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmcytID_CT = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Editar_Curva_y_Talla.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE CURVA Y TALLA] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_CT"
            Call CARGAR_DATAGRIDVIEW_7()
            Call ARMAR_DATAGRIDVIEW_7()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_7()
        End If
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_7()
        Dim I As Integer
        For I = 0 To Me.DGV7.RowCount - 1
            If Me.DGV7.Item(0, I).Value = vmcytID_CT Then
                Me.DGV7.Rows(I).Selected = True
                Me.DGV7.CurrentCell = Me.DGV7.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub Button29_Click(sender As Object, e As EventArgs) Handles Button29.Click 'AGREGAR
        Call VERIFICAR_ACCESOS("123") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Agregar_Curva_y_Talla.ShowDialog()
    End Sub
    'VISTA MAESTRO DE CLASIFICACION DE PERSONAL
    Private Sub DGV8_Click(sender As Object, e As EventArgs) Handles DGV8.Click
        If Me.DGV8.RowCount > 0 Then
            For I = 0 To Me.DGV8.RowCount - 1
                If DGV8.Rows(I).Selected = True Then
                    If Not IsDBNull(Me.DGV8.Item(0, I).Value) Then : vmcoID_C_PERSONAL = Me.DGV8.Item(0, I).Value : Else : vmcoID_C_PERSONAL = 0 : End If
                    If Not IsDBNull(Me.DGV8.Item(1, I).Value) Then : vmcoID_M_P = Me.DGV8.Item(1, I).Value : Else : vmcoID_M_P = 0 : End If
                    If Not IsDBNull(Me.DGV8.Item(2, I).Value) Then : vmcoID_C_PE = Me.DGV8.Item(2, I).Value : Else : vmcoID_C_PE = 0 : End If
                    If Not IsDBNull(Me.DGV8.Item(3, I).Value) Then : vmcoD_C_PE = Me.DGV8.Item(3, I).Value : Else : vmcoD_C_PE = "" : End If
                    If Not IsDBNull(Me.DGV8.Item(4, I).Value) Then : vmcoOBSERVACIONES = Me.DGV8.Item(4, I).Value : Else : vmcoOBSERVACIONES = "" : End If
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub Button30_Click(sender As Object, e As EventArgs) Handles Button30.Click   'ELIMINAR
        Call VERIFICAR_ACCESOS("128") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmcoID_C_PERSONAL = 0 Then
            MsgBox("Debe seleccionar el registro que desea Eliminar", vbInformation, "Mensaje del Sistema")
            Me.DGV8.Focus()
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Esta seguro de Eliminar el registro seleccionado?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Conexion.ABD_RRHH()
            Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
            Try
                MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE CLASIFICACION DE PERSONAL] WHERE ID_C_PERSONAL = " & CInt(vmcoID_C_PERSONAL) & ""
                MiSqlCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
            Call CARGAR_DATAGRIDVIEW_8()
            Call ARMAR_DATAGRIDVIEW_8()
        Else
            vmcoID_C_PERSONAL = 0
        End If
    End Sub
    Private Sub DGV8_DoubleClick(sender As Object, e As EventArgs) Handles DGV8.DoubleClick 'EDITAR
        Call VERIFICAR_ACCESOS("127") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmcoID_C_PERSONAL = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Editar_Clasificacion.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE CLASIFICACION DE PERSONAL] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_C_PERSONAL"
            Call CARGAR_DATAGRIDVIEW_8()
            Call ARMAR_DATAGRIDVIEW_8()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_8()
        End If
    End Sub
    Private Sub Button31_Click(sender As Object, e As EventArgs) Handles Button31.Click 'EDITAR
        Call VERIFICAR_ACCESOS("127") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmcoID_C_PERSONAL = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Editar_Clasificacion.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE CLASIFICACION DE PERSONAL] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_C_PERSONAL"
            Call CARGAR_DATAGRIDVIEW_8()
            Call ARMAR_DATAGRIDVIEW_8()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_8()
        End If
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_8()
        Dim I As Integer
        For I = 0 To Me.DGV8.RowCount - 1
            If Me.DGV8.Item(0, I).Value = vmcoID_C_PERSONAL Then
                Me.DGV8.Rows(I).Selected = True
                Me.DGV8.CurrentCell = Me.DGV8.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub Button32_Click(sender As Object, e As EventArgs) Handles Button32.Click 'AGREGAR
        Call VERIFICAR_ACCESOS("126") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Agregar_Clasificacion.ShowDialog()
    End Sub
    'VISTA MAESTRO DE CONDECORACIONES
    Private Sub DGV9_Click(sender As Object, e As EventArgs) Handles DGV9.Click
        If Me.DGV9.RowCount > 0 Then
            For I = 0 To Me.DGV9.RowCount - 1
                If DGV9.Rows(I).Selected = True Then
                    If Not IsDBNull(Me.DGV9.Item(0, I).Value) Then : vmcID_C_CONDE = Me.DGV9.Item(0, I).Value : Else : vmcID_C_CONDE = 0 : End If
                    If Not IsDBNull(Me.DGV9.Item(1, I).Value) Then : vmcD_C_CONDE = Me.DGV9.Item(1, I).Value : Else : vmcD_C_CONDE = "" : End If
                    If Not IsDBNull(Me.DGV9.Item(2, I).Value) Then : vmcID_CONDE = Me.DGV9.Item(2, I).Value : Else : vmcID_CONDE = 0 : End If
                    If Not IsDBNull(Me.DGV9.Item(3, I).Value) Then : vmcCODIGO = Me.DGV9.Item(3, I).Value : Else : vmcCODIGO = "" : End If
                    If Not IsDBNull(Me.DGV9.Item(4, I).Value) Then : vmcD_CONDE = Me.DGV9.Item(4, I).Value : Else : vmcD_CONDE = 0 : End If
                    If Not IsDBNull(Me.DGV9.Item(5, I).Value) Then : vmcID_M_CONDE = Me.DGV9.Item(5, I).Value : Else : vmcID_M_CONDE = 0 : End If
                    If Not IsDBNull(Me.DGV9.Item(6, I).Value) Then : vmcID_M_P = Me.DGV9.Item(6, I).Value : Else : vmcID_M_P = 0 : End If
                    If Not IsDBNull(Me.DGV9.Item(7, I).Value) Then : vmcCONDE_A = Me.DGV9.Item(7, I).Value : Else : vmcCONDE_A = "" : End If
                    If Not IsDBNull(Me.DGV9.Item(8, I).Value) Then : vmcMOTIVOS = Me.DGV9.Item(8, I).Value : Else : vmcMOTIVOS = "" : End If
                    If Not IsDBNull(Me.DGV9.Item(9, I).Value) Then : vmcOBSERVACIONES = Me.DGV9.Item(9, I).Value : Else : vmcOBSERVACIONES = "" : End If
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub Button33_Click(sender As Object, e As EventArgs) Handles Button33.Click   'ELIMINAR
        Call VERIFICAR_ACCESOS("131") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmcID_M_CONDE = 0 Then
            MsgBox("Debe seleccionar el registro que desea Eliminar", vbInformation, "Mensaje del Sistema")
            Me.DGV9.Focus()
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Esta seguro de Eliminar el registro seleccionado?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Conexion.ABD_RRHH()
            Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
            Try
                MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE CONDECORACIONES] WHERE ID_M_CONDE = " & CInt(vmcID_M_CONDE) & ""
                MiSqlCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
            Call CARGAR_DATAGRIDVIEW_9()
            Call ARMAR_DATAGRIDVIEW_9()
        Else
            vmcID_M_CONDE = 0
        End If
    End Sub
    Private Sub DGV9_DoubleClick(sender As Object, e As EventArgs) Handles DGV9.DoubleClick 'EDITAR
        Call VERIFICAR_ACCESOS("130") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmcID_M_CONDE = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Editar_Condecoraciones.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE CONDECORACIONES] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_M_CONDE"
            Call CARGAR_DATAGRIDVIEW_9()
            Call ARMAR_DATAGRIDVIEW_9()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_9()
        End If
    End Sub
    Private Sub Button34_Click(sender As Object, e As EventArgs) Handles Button34.Click 'EDITAR
        Call VERIFICAR_ACCESOS("130") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmcID_M_CONDE = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Editar_Condecoraciones.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE CONDECORACIONES] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_M_CONDE"
            Call CARGAR_DATAGRIDVIEW_9()
            Call ARMAR_DATAGRIDVIEW_9()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_9()
        End If
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_9()
        Dim I As Integer
        For I = 0 To Me.DGV9.RowCount - 1
            If Me.DGV9.Item(5, I).Value = vmcID_M_CONDE Then
                Me.DGV9.Rows(I).Selected = True
                Me.DGV9.CurrentCell = Me.DGV9.Item(5, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub Button35_Click(sender As Object, e As EventArgs) Handles Button35.Click 'AGREGAR
        Call VERIFICAR_ACCESOS("129") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Agregar_Condecoraciones.ShowDialog()
    End Sub
    'VISTA MAESTRO DE EVENTUALIDADES
    Private Sub DGV10_Click(sender As Object, e As EventArgs) Handles DGV10.Click
        If Me.DGV10.RowCount > 0 Then
            For I = 0 To Me.DGV10.RowCount - 1
                If DGV10.Rows(I).Selected = True Then
                    If Not IsDBNull(Me.DGV10.Item(0, I).Value) Then : vmeID_M_EV = Me.DGV10.Item(0, I).Value : Else : vmeID_M_EV = 0 : End If
                    If Not IsDBNull(Me.DGV10.Item(1, I).Value) Then : vmeID_M_P = Me.DGV10.Item(1, I).Value : Else : vmeID_M_P = 0 : End If
                    If Not IsDBNull(Me.DGV10.Item(2, I).Value) Then : vmeD_N1 = Me.DGV10.Item(2, I).Value : Else : vmeD_N1 = "" : End If
                    If Not IsDBNull(Me.DGV10.Item(3, I).Value) Then : vmeD_N2 = Me.DGV10.Item(3, I).Value : Else : vmeD_N2 = "" : End If
                    If Not IsDBNull(Me.DGV10.Item(4, I).Value) Then : vmeJEFE = Me.DGV10.Item(4, I).Value : Else : vmeJEFE = "" : End If
                    If Not IsDBNull(Me.DGV10.Item(5, I).Value) Then : vmeFEC_EVE = Me.DGV10.Item(5, I).Value : Else : vmeFEC_EVE = 0 : End If
                    If Not IsDBNull(Me.DGV10.Item(6, I).Value) Then : vmeID_CR_EVENTOS = Me.DGV10.Item(6, I).Value : Else : vmeID_CR_EVENTOS = 0 : End If
                    If Not IsDBNull(Me.DGV10.Item(7, I).Value) Then : vmeD_CR_EVENTOS = Me.DGV10.Item(7, I).Value : Else : vmeD_CR_EVENTOS = "" : End If
                    If Not IsDBNull(Me.DGV10.Item(8, I).Value) Then : vmeID_E = Me.DGV10.Item(8, I).Value : Else : vmeID_E = 0 : End If
                    If Not IsDBNull(Me.DGV10.Item(9, I).Value) Then : vmeD_E = Me.DGV10.Item(9, I).Value : Else : vmeD_E = "" : End If
                    If Not IsDBNull(Me.DGV10.Item(10, I).Value) Then : vmeEVENTUALIDAD = Me.DGV10.Item(10, I).Value : Else : vmeEVENTUALIDAD = "" : End If
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub Button36_Click(sender As Object, e As EventArgs) Handles Button36.Click   'ELIMINAR
        Call VERIFICAR_ACCESOS("134") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmeID_M_EV = 0 Then
            MsgBox("Debe seleccionar el registro que desea Eliminar", vbInformation, "Mensaje del Sistema")
            Me.DGV10.Focus()
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Esta seguro de Eliminar el registro seleccionado?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Conexion.ABD_RRHH()
            Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
            Try
                MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE EVENTUALIDADES] WHERE ID_M_EV = " & CInt(vmeID_M_EV) & ""
                MiSqlCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
            Call CARGAR_DATAGRIDVIEW_10()
            Call ARMAR_DATAGRIDVIEW_10()
        Else
            vmeID_M_EV = 0
        End If
    End Sub
    Private Sub DGV10_DoubleClick(sender As Object, e As EventArgs) Handles DGV10.DoubleClick   'EDITAR
        Call VERIFICAR_ACCESOS("133") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmeID_M_EV = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Editar_Eventualidades.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE EVENTUALIDADES] WHERE ID_M_P = " & vID_M_P & " ORDER BY FEC_EVE DESC"
            Call CARGAR_DATAGRIDVIEW_10()
            Call ARMAR_DATAGRIDVIEW_10()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_10()
        End If
    End Sub
    Private Sub Button37_Click(sender As Object, e As EventArgs) Handles Button37.Click 'EDITAR
        Call VERIFICAR_ACCESOS("133") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmeID_M_EV = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Editar_Eventualidades.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE EVENTUALIDADES] WHERE ID_M_P = " & vID_M_P & " ORDER BY FEC_EVE DESC"
            Call CARGAR_DATAGRIDVIEW_10()
            Call ARMAR_DATAGRIDVIEW_10()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_10()
        End If
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_10()
        Dim I As Integer
        For I = 0 To Me.DGV10.RowCount - 1
            If Me.DGV10.Item(0, I).Value = vmeID_M_EV Then
                Me.DGV10.Rows(I).Selected = True
                Me.DGV10.CurrentCell = Me.DGV10.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub Button38_Click(sender As Object, e As EventArgs) Handles Button38.Click 'AGREGAR
        Call VERIFICAR_ACCESOS("132") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Agregar_Eventualidades.ShowDialog()
    End Sub
    'VISTA MAESTRO DE NOMINAS Y SALARIOS
    Private Sub DGV11_Click(sender As Object, e As EventArgs) Handles DGV11.Click
        If Me.DGV11.RowCount > 0 Then
            For I = 0 To Me.DGV11.RowCount - 1
                If DGV11.Rows(I).Selected = True Then
                    If Not IsDBNull(Me.DGV11.Item(0, I).Value) Then : vmnsID_M_NYS = Me.DGV11.Item(0, I).Value : Else : vmnsID_M_NYS = 0 : End If
                    If Not IsDBNull(Me.DGV11.Item(1, I).Value) Then : vmnsID_M_P = Me.DGV11.Item(1, I).Value : Else : vmnsID_M_P = 0 : End If
                    If Not IsDBNull(Me.DGV11.Item(2, I).Value) Then : vmnsID_MOVIMIENTO_N = Me.DGV11.Item(2, I).Value : Else : vmnsID_MOVIMIENTO_N = 0 : End If
                    If Not IsDBNull(Me.DGV11.Item(3, I).Value) Then : vmnsD_MOVIMIENTO_N = Me.DGV11.Item(3, I).Value : Else : vmnsD_MOVIMIENTO_N = "" : End If
                    If Not IsDBNull(Me.DGV11.Item(4, I).Value) Then : vmnsF_MOVIMIENTO_N = Me.DGV11.Item(4, I).Value : Else : vmnsF_MOVIMIENTO_N = "" : End If
                    If Not IsDBNull(Me.DGV11.Item(5, I).Value) Then : vmnsID_NOMINA = Me.DGV11.Item(5, I).Value : Else : vmnsID_NOMINA = 0 : End If
                    If Not IsDBNull(Me.DGV11.Item(6, I).Value) Then : vmnsD_NOMINA = Me.DGV11.Item(6, I).Value : Else : vmnsD_NOMINA = 0 : End If
                    If Not IsDBNull(Me.DGV11.Item(7, I).Value) Then : vmnsPROPUESTA = Me.DGV11.Item(7, I).Value : Else : vmnsPROPUESTA = 0 : End If
                    If Not IsDBNull(Me.DGV11.Item(8, I).Value) Then : vmnsSALARIO = Me.DGV11.Item(8, I).Value : Else : vmnsSALARIO = 0 : End If
                    If Not IsDBNull(Me.DGV11.Item(9, I).Value) Then : vmnsSAL_MINIMO = Me.DGV11.Item(9, I).Value : Else : vmnsSAL_MINIMO = 0 : End If
                    If Not IsDBNull(Me.DGV11.Item(10, I).Value) Then : vmnsOBSERVACIONES = Me.DGV11.Item(10, I).Value : Else : vmnsOBSERVACIONES = 0 : End If
                    If Not IsDBNull(Me.DGV11.Item(11, I).Value) Then : vmnsACTIVO = Me.DGV11.Item(11, I).Value : Else : vmnsACTIVO = False : End If
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub Button39_Click(sender As Object, e As EventArgs) Handles Button39.Click   'ELIMINAR
        Call VERIFICAR_ACCESOS("137") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmnsID_M_NYS = 0 Then
            MsgBox("Debe seleccionar el registro que desea Eliminar", vbInformation, "Mensaje del Sistema")
            Me.DGV11.Focus()
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Esta seguro de Eliminar el registro seleccionado?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Conexion.ABD_RRHH()
            Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
            Try
                MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE NOMINAS Y SALARIOS] WHERE ID_M_NYS = " & CInt(vmnsID_M_NYS) & ""
                MiSqlCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
            Call CARGAR_DATAGRIDVIEW_11()
            Call ARMAR_DATAGRIDVIEW_11()
        Else
            vmnsID_M_NYS = 0
        End If
    End Sub
    Private Sub DGV11_DoubleClick(sender As Object, e As EventArgs) Handles DGV11.DoubleClick   'EDITAR
        Call VERIFICAR_ACCESOS("136") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmnsID_M_NYS = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Editar_NyS.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE NOMINAS Y SALARIOS] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_NOMINA, F_MOVIMIENTO_N DESC"
            Call CARGAR_DATAGRIDVIEW_11()
            Call ARMAR_DATAGRIDVIEW_11()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_11()
        End If
    End Sub
    Private Sub Button40_Click(sender As Object, e As EventArgs) Handles Button40.Click 'EDITAR
        Call VERIFICAR_ACCESOS("136") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmnsID_M_NYS = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Editar_NyS.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE NOMINAS Y SALARIOS] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_NOMINA, F_MOVIMIENTO_N DESC"
            Call CARGAR_DATAGRIDVIEW_11()
            Call ARMAR_DATAGRIDVIEW_11()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_11()
        End If
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_11()
        Dim I As Integer
        For I = 0 To Me.DGV11.RowCount - 1
            If Me.DGV11.Item(0, I).Value = vmnsID_M_NYS Then
                Me.DGV11.Rows(I).Selected = True
                Me.DGV11.CurrentCell = Me.DGV11.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub Button41_Click(sender As Object, e As EventArgs) Handles Button41.Click 'AGREGAR
        Call VERIFICAR_ACCESOS("135") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Agregar_NyS.ShowDialog()
    End Sub
    'VISTA MAESTRO DE DOCUMENTOS DIGITALES
    Private Sub DGV12_Click(sender As Object, e As EventArgs) Handles DGV12.Click
        If Me.DGV12.RowCount > 0 Then
            For I = 0 To Me.DGV12.RowCount - 1
                If DGV12.Rows(I).Selected = True Then
                    If Not IsDBNull(Me.DGV12.Item(0, I).Value) Then : vmdID_M_DD = Me.DGV12.Item(0, I).Value : Else : vmdID_M_DD = 0 : End If
                    If Not IsDBNull(Me.DGV12.Item(1, I).Value) Then : vmdID_M_P = Me.DGV12.Item(1, I).Value : Else : vmdID_M_P = 0 : End If
                    If Not IsDBNull(Me.DGV12.Item(2, I).Value) Then : vmdID_D1 = Me.DGV12.Item(2, I).Value : Else : vmdID_D1 = 0 : End If
                    If Not IsDBNull(Me.DGV12.Item(3, I).Value) Then : vmdDIRECTORIO = Me.DGV12.Item(3, I).Value : Else : vmdDIRECTORIO = "" : End If
                    If Not IsDBNull(Me.DGV12.Item(4, I).Value) Then : vmdID_T_DOC = Me.DGV12.Item(4, I).Value : Else : vmdID_T_DOC = 0 : End If
                    If Not IsDBNull(Me.DGV12.Item(5, I).Value) Then : vmdD_T_DOC = Me.DGV12.Item(5, I).Value : Else : vmdD_T_DOC = "" : End If
                    If Not IsDBNull(Me.DGV12.Item(6, I).Value) Then : vmdNOMBRE = Me.DGV12.Item(6, I).Value : Else : vmdNOMBRE = "" : End If
                    If Not IsDBNull(Me.DGV12.Item(7, I).Value) Then : vmdOBSERVACION = Me.DGV12.Item(7, I).Value : Else : vmdOBSERVACION = "" : End If
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub Button42_Click(sender As Object, e As EventArgs) Handles Button42.Click   'ELIMINAR
        Call VERIFICAR_ACCESOS("144") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmdID_M_DD = 0 Then
            MsgBox("Debe seleccionar el registro que desea Eliminar", vbInformation, "Mensaje del Sistema")
            Me.DGV12.Focus()
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Esta seguro de Eliminar el registro seleccionado?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Conexion.ABD_RRHH()
            Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
            Try

                Dim FILE As String = vmdDIRECTORIO & "\" & vmdNOMBRE

                My.Computer.FileSystem.DeleteFile(FILE, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin, FileIO.UICancelOption.ThrowException)

                MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE DOCUMENTOS DIGITALES] WHERE ID_M_DD = " & CInt(vmdID_M_DD) & ""
                MiSqlCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
            Call CARGAR_DATAGRIDVIEW_12()
            Call ARMAR_DATAGRIDVIEW_12()
        Else
            vmdID_M_DD = 0
        End If
    End Sub
    Private Sub DGV12_DoubleClick(sender As Object, e As EventArgs) Handles DGV12.DoubleClick   'EDITAR
        Call VERIFICAR_ACCESOS("143") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmdID_M_DD = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Editar_Documentos.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE DOCUMENTOS DIGITALES] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_T_DOC"
            Call CARGAR_DATAGRIDVIEW_12()
            Call ARMAR_DATAGRIDVIEW_12()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_12()
        End If
    End Sub
    Private Sub Button43_Click(sender As Object, e As EventArgs) Handles Button43.Click 'EDITAR
        Call VERIFICAR_ACCESOS("143") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmdID_M_DD = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Editar_Documentos.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE DOCUMENTOS DIGITALES] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_T_DOC"
            Call CARGAR_DATAGRIDVIEW_12()
            Call ARMAR_DATAGRIDVIEW_12()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_12()
        End If
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_12()
        Dim I As Integer
        For I = 0 To Me.DGV12.RowCount - 1
            If Me.DGV12.Item(0, I).Value = vmdID_M_DD Then
                Me.DGV12.Rows(I).Selected = True
                Me.DGV12.CurrentCell = Me.DGV12.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub Button44_Click(sender As Object, e As EventArgs) Handles Button44.Click 'AGREGAR
        Call VERIFICAR_ACCESOS("142") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Agregar_Documentos.ShowDialog()
    End Sub
    Private Sub Button51_Click(sender As Object, e As EventArgs) Handles Button51.Click 'VISUALIZAR
        Call VERIFICAR_ACCESOS("141") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmdID_M_DD = 0 Then
            MsgBox("Debe seleccionar el registro que desea Visualizar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Visor_Documentos.ShowDialog()
    End Sub
    'VISTA MAESTRO DE SITUACION
    Private Sub DGV13_Click(sender As Object, e As EventArgs) Handles DGV13.Click
        If Me.DGV13.RowCount > 0 Then
            For I = 0 To Me.DGV13.RowCount - 1
                If DGV13.Rows(I).Selected = True Then
                    If Not IsDBNull(Me.DGV13.Item(1, I).Value) Then : msNIVEL1 = Me.DGV13.Item(1, I).Value : Else : msNIVEL1 = "" : End If
                    If Not IsDBNull(Me.DGV13.Item(2, I).Value) Then : msNIVEL2 = Me.DGV13.Item(1, I).Value : Else : msNIVEL2 = "" : End If
                    If Not IsDBNull(Me.DGV13.Item(3, I).Value) Then : msNIVEL3 = Me.DGV13.Item(3, I).Value : Else : msNIVEL3 = "" : End If
                    If Not IsDBNull(Me.DGV13.Item(4, I).Value) Then : msNIVEL4 = Me.DGV13.Item(4, I).Value : Else : msNIVEL4 = "" : End If
                    If Not IsDBNull(Me.DGV13.Item(5, I).Value) Then : msNIVEL5 = Me.DGV13.Item(5, I).Value : Else : msNIVEL5 = "" : End If
                    If Not IsDBNull(Me.DGV13.Item(6, I).Value) Then : msORD = Me.DGV13.Item(6, I).Value : Else : msORD = "" : End If
                    If Not IsDBNull(Me.DGV13.Item(7, I).Value) Then : msID_M_A = Me.DGV13.Item(7, I).Value : Else : msID_M_A = 0 : End If
                    If Not IsDBNull(Me.DGV13.Item(8, I).Value) Then : msID_M_P = Me.DGV13.Item(8, I).Value : Else : msID_M_P = 0 : End If
                    If Not IsDBNull(Me.DGV13.Item(9, I).Value) Then : msCODIGO = Me.DGV13.Item(9, I).Value : Else : msCODIGO = "" : End If
                    If Not IsDBNull(Me.DGV13.Item(10, I).Value) Then : msDIAL = Me.DGV13.Item(10, I).Value : Else : msDIAL = "" : End If
                    If Not IsDBNull(Me.DGV13.Item(11, I).Value) Then : msDIA = Me.DGV13.Item(11, I).Value : Else : msDIA = "" : End If
                    If Not IsDBNull(Me.DGV13.Item(12, I).Value) Then : msID_T_SIT_P = Me.DGV13.Item(12, I).Value : Else : msID_T_SIT_P = 0 : End If
                    If Not IsDBNull(Me.DGV13.Item(13, I).Value) Then : msD_T_SIT_P = Me.DGV13.Item(13, I).Value : Else : msD_T_SIT_P = "" : End If
                    If Not IsDBNull(Me.DGV13.Item(14, I).Value) Then : msID_SIT_P = Me.DGV13.Item(14, I).Value : Else : msID_SIT_P = 0 : End If
                    If Not IsDBNull(Me.DGV13.Item(15, I).Value) Then : msD_SIT_P = Me.DGV13.Item(15, I).Value : Else : msD_SIT_P = "" : End If
                    If Not IsDBNull(Me.DGV13.Item(16, I).Value) Then : msSIGLA = Me.DGV13.Item(16, I).Value : Else : msSIGLA = "" : End If
                    If Not IsDBNull(Me.DGV13.Item(17, I).Value) Then : msVALOR_SIT = Me.DGV13.Item(17, I).Value : Else : msVALOR_SIT = "" : End If
                    If Not IsDBNull(Me.DGV13.Item(18, I).Value) Then : msVALOR_DIA = Me.DGV13.Item(18, I).Value : Else : msVALOR_DIA = "" : End If
                    If Not IsDBNull(Me.DGV13.Item(19, I).Value) Then : msOBSERVACIONES = Me.DGV13.Item(19, I).Value : Else : msOBSERVACIONES = "" : End If
                    If Not IsDBNull(Me.DGV13.Item(20, I).Value) Then : msDETALLE_1 = Me.DGV13.Item(20, I).Value : Else : msDETALLE_1 = "" : End If
                    If Not IsDBNull(Me.DGV13.Item(21, I).Value) Then : msDETALLE_2 = Me.DGV13.Item(21, I).Value : Else : msDETALLE_2 = "" : End If
                    If Not IsDBNull(Me.DGV13.Item(22, I).Value) Then : msDETALLE_3 = Me.DGV13.Item(22, I).Value : Else : msDETALLE_3 = "" : End If
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub DateTimePicker7_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker7.ValueChanged
        FECHA01 = Me.DateTimePicker7.Text
        FECHA01 = Mid(FECHA01, 1, 10)
        FECHA02 = Me.DateTimePicker6.Text
        FECHA02 = Mid(FECHA02, 1, 10)
        If Len(Me.TextBox2.Text) = 0 Or (Me.TextBox2.Text = "0000000000") Then
            Exit Sub
        End If
        CADENAsql = "SET LANGUAGE ESPAÑOL Select ID, NIVEL1, NIVEL2, NIVEL3, NIVEL4, NIVEL5, ORD, ID_M_A, ID_M_P, CODIGO, UPPER(DATENAME (dw, DIA)) AS DIALETRAS, DIA, ID_T_SIT_P, D_T_SIT_P, ID_SIT_P, D_SIT_P, SIGLA, VALOR_SIT, VALOR_DIA, OBSERVACIONES, DETALLE_1, DETALLE_2, DETALLE_3 from [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE CODIGO = '" & Me.TextBox2.Text & "' AND DIA BETWEEN '" & FECHA01 & "' AND '" & FECHA02 & "' ORDER BY CODIGO, DIA SET LANGUAGE ENGLISH"
        Call CARGAR_DATAGRIDVIEW_13()
        Call ARMAR_DATAGRIDVIEW_13()
    End Sub
    Private Sub DateTimePicker6_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker6.ValueChanged
        FECHA01 = Me.DateTimePicker7.Text
        FECHA01 = Mid(FECHA01, 1, 10)
        FECHA02 = Me.DateTimePicker6.Text
        FECHA02 = Mid(FECHA02, 1, 10)
        If Len(Me.TextBox2.Text) = 0 Or (Me.TextBox2.Text = "0000000000") Then
            Exit Sub
        End If
        CADENAsql = "SET LANGUAGE ESPAÑOL Select ID, NIVEL1, NIVEL2, NIVEL3, NIVEL4, NIVEL5, ORD, ID_M_A, ID_M_P, CODIGO, UPPER(DATENAME (dw, DIA)) AS DIALETRAS, DIA, ID_T_SIT_P, D_T_SIT_P, ID_SIT_P, D_SIT_P, SIGLA, VALOR_SIT, VALOR_DIA, OBSERVACIONES, DETALLE_1, DETALLE_2, DETALLE_3 from [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE CODIGO = '" & Me.TextBox2.Text & "' AND DIA BETWEEN '" & FECHA01 & "' AND '" & FECHA02 & "' ORDER BY CODIGO, DIA SET LANGUAGE ENGLISH"
        Call CARGAR_DATAGRIDVIEW_13()
        Call ARMAR_DATAGRIDVIEW_13()
    End Sub
    Private Sub Button46_Click(sender As Object, e As EventArgs)  'EDITAR
        If msID_M_A = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        'Frm_Editar_Anexos.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE ANEXOS] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_T_ANEXO, ID_ANEXO"
            Call CARGAR_DATAGRIDVIEW_13()
            Call ARMAR_DATAGRIDVIEW_13()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_13()
        End If
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_13()
        Dim I As Integer
        For I = 0 To Me.DGV13.RowCount - 1
            If Me.DGV13.Item(7, I).Value = msID_M_A Then
                Me.DGV13.Rows(I).Selected = True
                Me.DGV13.CurrentCell = Me.DGV13.Item(7, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub Button47_Click(sender As Object, e As EventArgs)  'AGREGAR
        'Frm_Agregar_Anexos.ShowDialog()
    End Sub
    'VISTA HISTORICO DE MOVIMIENTO
    Private Sub DGV14_Click(sender As Object, e As EventArgs) Handles DGV14.Click
        If Me.DGV14.RowCount > 0 Then
            For I = 0 To Me.DGV14.RowCount - 1
                If DGV14.Rows(I).Selected = True Then
                    If Not IsDBNull(Me.DGV14.Item(0, I).Value) Then : hmID_HIST_M = Me.DGV14.Item(0, I).Value : Else : hmID_HIST_M = 0 : End If
                    If Not IsDBNull(Me.DGV14.Item(1, I).Value) Then : hmFEC_ING_PAME = Me.DGV14.Item(1, I).Value : Else : hmFEC_ING_PAME = "" : End If
                    If Not IsDBNull(Me.DGV14.Item(2, I).Value) Then : hmID_M_P = Me.DGV14.Item(2, I).Value : Else : hmID_M_P = 0 : End If
                    If Not IsDBNull(Me.DGV14.Item(3, I).Value) Then : hmCODIGO = Me.DGV14.Item(3, I).Value : Else : hmCODIGO = "" : End If
                    If Not IsDBNull(Me.DGV14.Item(4, I).Value) Then : hmAN = Me.DGV14.Item(4, I).Value : Else : hmAN = "" : End If
                    If Not IsDBNull(Me.DGV14.Item(5, I).Value) Then : hmFEC_MOV = Me.DGV14.Item(5, I).Value : Else : hmFEC_MOV = "" : End If
                    If Not IsDBNull(Me.DGV14.Item(6, I).Value) Then : hmFEC_APLICA = Me.DGV14.Item(6, I).Value : Else : hmFEC_APLICA = "" : End If

                    If Not IsDBNull(Me.DGV14.Item(7, I).Value) Then : hmID_T_M_EXP = Me.DGV14.Item(7, I).Value : Else : hmID_T_M_EXP = 0 : End If
                    If Not IsDBNull(Me.DGV14.Item(8, I).Value) Then : hmD_T_M_EXP = Me.DGV14.Item(8, I).Value : Else : hmD_T_M_EXP = "" : End If
                    If Not IsDBNull(Me.DGV14.Item(9, I).Value) Then : hmID_ST_M_L_EXP = Me.DGV14.Item(9, I).Value : Else : hmID_ST_M_L_EXP = 0 : End If
                    If Not IsDBNull(Me.DGV14.Item(10, I).Value) Then : hmD_ST_M_L_EXP = Me.DGV14.Item(10, I).Value : Else : hmD_ST_M_L_EXP = "" : End If
                    If Not IsDBNull(Me.DGV14.Item(11, I).Value) Then : hmID_ST_M_R_EXP = Me.DGV14.Item(11, I).Value : Else : hmID_ST_M_R_EXP = 0 : End If
                    If Not IsDBNull(Me.DGV14.Item(12, I).Value) Then : hmD_ST_M_R_EXP = Me.DGV14.Item(12, I).Value : Else : hmD_ST_M_R_EXP = "" : End If
                    If Not IsDBNull(Me.DGV14.Item(13, I).Value) Then : hmD_CARGO_ES = Me.DGV14.Item(13, I).Value : Else : hmD_CARGO_ES = "" : End If
                    If Not IsDBNull(Me.DGV14.Item(14, I).Value) Then : hmD_GP = Me.DGV14.Item(14, I).Value : Else : hmD_GP = "" : End If
                    If Not IsDBNull(Me.DGV14.Item(15, I).Value) Then : hmD_GR = Me.DGV14.Item(15, I).Value : Else : hmD_GR = "" : End If
                    If Not IsDBNull(Me.DGV14.Item(16, I).Value) Then : hmD_N1 = Me.DGV14.Item(16, I).Value : Else : hmD_N1 = "" : End If
                    If Not IsDBNull(Me.DGV14.Item(17, I).Value) Then : hmD_N2 = Me.DGV14.Item(17, I).Value : Else : hmD_N2 = "" : End If
                    If Not IsDBNull(Me.DGV14.Item(18, I).Value) Then : hmD_N3 = Me.DGV14.Item(18, I).Value : Else : hmD_N3 = "" : End If
                    If Not IsDBNull(Me.DGV14.Item(19, I).Value) Then : hmD_N4 = Me.DGV14.Item(19, I).Value : Else : hmD_N4 = "" : End If
                    If Not IsDBNull(Me.DGV14.Item(20, I).Value) Then : hmD_N5 = Me.DGV14.Item(20, I).Value : Else : hmD_N5 = "" : End If
                    If Not IsDBNull(Me.DGV14.Item(21, I).Value) Then : hmN_ORDEN = Me.DGV14.Item(21, I).Value : Else : hmN_ORDEN = "" : End If
                    If Not IsDBNull(Me.DGV14.Item(22, I).Value) Then : hmD_NOMINA = Me.DGV14.Item(22, I).Value : Else : hmD_NOMINA = "" : End If
                    If Not IsDBNull(Me.DGV14.Item(23, I).Value) Then : hmSALARIO = Me.DGV14.Item(23, I).Value : Else : hmSALARIO = "" : End If
                    If Not IsDBNull(Me.DGV14.Item(24, I).Value) Then : hmDETALLE = Me.DGV14.Item(24, I).Value : Else : hmDETALLE = "" : End If
                    If Not IsDBNull(Me.DGV14.Item(25, I).Value) Then : hmMIXTA = Me.DGV14.Item(25, I).Value : Else : hmMIXTA = "" : End If
                    If Not IsDBNull(Me.DGV14.Item(26, I).Value) Then : hmMILITAR = Me.DGV14.Item(26, I).Value : Else : hmMILITAR = "" : End If
                    If Not IsDBNull(Me.DGV14.Item(27, I).Value) Then : hmPAME = Me.DGV14.Item(27, I).Value : Else : hmPAME = "" : End If
                    If Not IsDBNull(Me.DGV14.Item(28, I).Value) Then : hmAPLICAR = Me.DGV14.Item(28, I).Value : Else : hmAPLICAR = "" : End If
                    If Not IsDBNull(Me.DGV14.Item(29, I).Value) Then : hmESTADISTICO = Me.DGV14.Item(29, I).Value : Else : hmESTADISTICO = "" : End If
                    If Not IsDBNull(Me.DGV14.Item(30, I).Value) Then : hmUSUARIO_ACT = Me.DGV14.Item(30, I).Value : Else : hmUSUARIO_ACT = "" : End If
                    If Not IsDBNull(Me.DGV14.Item(31, I).Value) Then : hmFECHA_ACT = Me.DGV14.Item(31, I).Value : Else : hmFECHA_ACT = "" : End If
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub Frm_Expedientes_Activos_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        EXPEDIENTES = False
        vID_M_P = 0 : vID_M_C = 0
    End Sub
    Private Sub Frm_Expedientes_Activos_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        EXPEDIENTES = False
        vID_M_P = 0 : vID_M_C = 0
    End Sub
    Private Sub Frm_Expedientes_Activos_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        EXPEDIENTES = False
        vID_M_P = 0 : vID_M_C = 0
    End Sub
    Private Sub Frm_Expedientes_Activos_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        EXPEDIENTES = False
        vID_M_P = 0 : vID_M_C = 0
    End Sub
    Private Sub DateTimePicker4_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker4.ValueChanged
        FECHA01 = "Convert(VARCHAR(10), Convert(Date, '" + Me.DateTimePicker4.Value + "', 105), 23)"
        FECHA02 = "Convert(VARCHAR(10), Convert(Date, '" + Me.DateTimePicker5.Value + "', 105), 23)"
        CADENAsql = "SET LANGUAGE Español SELECT [IDEMPRESA], [RELOJ], [ID_EMPLEADO], DATENAME (dw, DIA) AS DIAL, [FECHA_MARCA], [HORA_MARCA], [FECHA_CARGA], [ARCHIVO], [NUMERO], [PROGRAMADO], [OBSERVACION], [TIPO_MARCA] FROM [dbo].[VISTA MAESTRO DE MARCAS] WHERE (ID_EMPLEADO = '" & Me.TextBox2.Text & "') AND (FECHA_MARCA BETWEEN " & FECHA01 & " AND " & FECHA02 & ")  ORDER BY FECHA_MARCA, HORA_MARCA SET LANGUAGE ENGLISH"
        Call CARGAR_DATAGRIDVIEW_15()
        Call ARMAR_DATAGRIDVIEW_15()
        Me.TextBox35.Text = vID_M_P
    End Sub
    Private Sub DateTimePicker5_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker5.ValueChanged
        FECHA01 = "Convert(VARCHAR(10), Convert(Date, '" + Me.DateTimePicker4.Value + "', 105), 23)"
        FECHA02 = "Convert(VARCHAR(10), Convert(Date, '" + Me.DateTimePicker5.Value + "', 105), 23)"
        CADENAsql = "SET LANGUAGE Español SELECT [IDEMPRESA], [RELOJ], [ID_EMPLEADO], DATENAME (dw, DIA) AS DIAL, [FECHA_MARCA], [HORA_MARCA], [FECHA_CARGA], [ARCHIVO], [NUMERO], [PROGRAMADO], [OBSERVACION], [TIPO_MARCA] FROM [dbo].[VISTA MAESTRO DE MARCAS] WHERE (ID_EMPLEADO = '" & Me.TextBox2.Text & "') AND (FECHA_MARCA BETWEEN " & FECHA01 & " AND " & FECHA02 & ")  ORDER BY FECHA_MARCA, HORA_MARCA SET LANGUAGE ENGLISH"
        Call CARGAR_DATAGRIDVIEW_15()
        Call ARMAR_DATAGRIDVIEW_15()
        Me.TextBox35.Text = vID_M_P
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Call VERIFICAR_ACCESOS("076") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Expedientes_Imprimir.ShowDialog()
    End Sub
    Private Sub DGV15_RowPrePaint(sender As Object, e As DataGridViewRowPrePaintEventArgs) Handles DGV15.RowPrePaint
        Dim style As New DataGridViewCellStyle
        style.Font = New Font(DGV15.Font, FontStyle.Bold)
        For Each Row As DataGridViewRow In DGV15.Rows
            If Row.Cells(3).Value = "Sábado" Or Row.Cells(3).Value = "Domingo" Then
                Row.DefaultCellStyle.ForeColor = Color.Red
            End If
            If Row.Cells(3).Value = "Lunes" Then
                Row.DefaultCellStyle.ForeColor = Color.Blue
            End If
        Next
    End Sub
    Private Sub ComboBox26_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox26.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    'VISTA MAESTRO DE HORARIOS
    Private Sub DGV16_Click(sender As Object, e As EventArgs) Handles DGV16.Click
        If Me.DGV16.RowCount > 0 Then
            For I = 0 To Me.DGV16.RowCount - 1
                If DGV16.Rows(I).Selected = True Then
                    If Not IsDBNull(Me.DGV16.Item(0, I).Value) Then : vmhiID_H = Me.DGV16.Item(0, I).Value : Else : vmhiID_H = 0 : End If
                    If Not IsDBNull(Me.DGV16.Item(1, I).Value) Then : vmhiID_M_P = Me.DGV16.Item(1, I).Value : Else : vmhiID_M_P = 0 : End If
                    If Not IsDBNull(Me.DGV16.Item(2, I).Value) Then : vmhiID_T_HORARIO = Me.DGV16.Item(2, I).Value : Else : vmhiID_T_HORARIO = 0 : End If
                    If Not IsDBNull(Me.DGV16.Item(3, I).Value) Then : vmhiD_T_HORARIO = Me.DGV16.Item(3, I).Value : Else : vmhiD_T_HORARIO = "" : End If
                    If Not IsDBNull(Me.DGV16.Item(4, I).Value) Then : vmhiHORA_E = Me.DGV16.Item(4, I).Value.ToString : Else : vmhiHORA_E = "" : End If
                    If Not IsDBNull(Me.DGV16.Item(5, I).Value) Then : vmhiHORA_S = Me.DGV16.Item(5, I).Value.ToString : Else : vmhiHORA_S = "" : End If
                    If Not IsDBNull(Me.DGV16.Item(6, I).Value) Then : vmhiCHRT = Me.DGV16.Item(6, I).Value : Else : vmhiCHRT = 0 : End If
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub DGV16_DoubleClick(sender As Object, e As EventArgs) Handles DGV16.DoubleClick
        Call VERIFICAR_ACCESOS("139") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmhiID_H = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Editar_Horarios.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HORARIOS] WHERE ID_M_P = " & vID_M_P & "ORDER BY HORA_E"
            Call CARGAR_DATAGRIDVIEW_16()
            Call ARMAR_DATAGRIDVIEW_16()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_16()
        End If
    End Sub
    Private Sub Button52_Click(sender As Object, e As EventArgs) Handles Button52.Click   'ELIMINAR
        Call VERIFICAR_ACCESOS("140") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmhiID_H = 0 Then
            MsgBox("Debe seleccionar el registro que desea Eliminar", vbInformation, "Mensaje del Sistema")
            Me.DGV16.Focus()
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Esta seguro de Eliminar el registro seleccionado?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Conexion.ABD_RRHH()
            Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
            Try
                MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE HORARIOS] WHERE ID_H = " & CInt(vmhiID_H) & ""
                MiSqlCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
            Call CARGAR_DATAGRIDVIEW_16()
            Call ARMAR_DATAGRIDVIEW_16()
        Else
            vmhiID_H = 0
        End If
    End Sub
    Private Sub Button53_Click(sender As Object, e As EventArgs) Handles Button53.Click 'EDITAR
        Call VERIFICAR_ACCESOS("139") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmhiID_H = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Editar_Horarios.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HORARIOS] WHERE ID_M_P = " & vID_M_P & "ORDER BY HORA_E"
            Call CARGAR_DATAGRIDVIEW_16()
            Call ARMAR_DATAGRIDVIEW_16()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_16()
        End If
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_16()
        Dim I As Integer
        For I = 0 To Me.DGV16.RowCount - 1
            If Me.DGV16.Item(0, I).Value = vmhiID_H Then
                Me.DGV16.Rows(I).Selected = True
                Me.DGV16.CurrentCell = Me.DGV16.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub Button54_Click(sender As Object, e As EventArgs) Handles Button54.Click 'AGREGAR
        Call VERIFICAR_ACCESOS("138") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Agregar_Horarios.ShowDialog()
    End Sub

    'VISTA MAESTRO DE SUBSIDIOS
    Private Sub DGV17_Click(sender As Object, e As EventArgs) Handles DGV17.Click
        If Me.DGV17.RowCount > 0 Then
            For I = 0 To Me.DGV17.RowCount - 1
                If DGV17.Rows(I).Selected = True Then
                    If Not IsDBNull(Me.DGV17.Item(0, I).Value) Then : vmssID_SUBS = Me.DGV17.Item(0, I).Value : Else : vmssID_SUBS = 0 : End If
                    If Not IsDBNull(Me.DGV17.Item(1, I).Value) Then : vmssID_M_P = Me.DGV17.Item(1, I).Value : Else : vmssID_M_P = 0 : End If
                    If Not IsDBNull(Me.DGV17.Item(2, I).Value) Then : vmssNO_EXPEDIENTE = Me.DGV17.Item(2, I).Value : Else : vmssNO_EXPEDIENTE = "" : End If
                    If Not IsDBNull(Me.DGV17.Item(3, I).Value) Then : vmssNO_BOLETA = Me.DGV17.Item(3, I).Value : Else : vmssNO_BOLETA = "" : End If
                    If Not IsDBNull(Me.DGV17.Item(4, I).Value) Then : vmssID_T_SUBS = Me.DGV17.Item(4, I).Value : Else : vmssID_T_SUBS = 0 : End If
                    If Not IsDBNull(Me.DGV17.Item(5, I).Value) Then : vmssD_T_SUBS = Me.DGV17.Item(5, I).Value : Else : vmssD_T_SUBS = "" : End If
                    If Not IsDBNull(Me.DGV17.Item(6, I).Value) Then : vmssN_ORDEN = Me.DGV17.Item(6, I).Value : Else : vmssN_ORDEN = "" : End If
                    If Not IsDBNull(Me.DGV17.Item(7, I).Value) Then : vmssFECHA_I = Me.DGV17.Item(7, I).Value : Else : vmssFECHA_I = "" : End If
                    If Not IsDBNull(Me.DGV17.Item(8, I).Value) Then : vmssFECHA_F = Me.DGV17.Item(8, I).Value : Else : vmssFECHA_F = "" : End If
                    If Not IsDBNull(Me.DGV17.Item(9, I).Value) Then : vmssCANT_DIAS = Me.DGV17.Item(9, I).Value : Else : vmssCANT_DIAS = "" : End If
                    If Not IsDBNull(Me.DGV17.Item(10, I).Value) Then : vmssFECHA_PARTO = Me.DGV17.Item(10, I).Value : Else : vmssFECHA_PARTO = "" : End If
                    If Not IsDBNull(Me.DGV17.Item(11, I).Value) Then : vmssFECHA_PARTO_PROB = Me.DGV17.Item(11, I).Value : Else : vmssFECHA_PARTO_PROB = "" : End If
                    If Not IsDBNull(Me.DGV17.Item(12, I).Value) Then : vmssFECHA_ACC_LAB = Me.DGV17.Item(12, I).Value : Else : vmssFECHA_ACC_LAB = "" : End If
                    If Not IsDBNull(Me.DGV17.Item(13, I).Value) Then : vmssFECHA_DECLARACION = Me.DGV17.Item(13, I).Value : Else : vmssFECHA_DECLARACION = "" : End If
                    If Not IsDBNull(Me.DGV17.Item(14, I).Value) Then : vmssVALOR_X_DIA = Me.DGV17.Item(14, I).Value : Else : vmssVALOR_X_DIA = "" : End If
                    If Not IsDBNull(Me.DGV17.Item(15, I).Value) Then : vmssVALOR_TOTAL = Me.DGV17.Item(15, I).Value : Else : vmssVALOR_TOTAL = "" : End If
                    If Not IsDBNull(Me.DGV17.Item(16, I).Value) Then : vmssMEDICO = Me.DGV17.Item(16, I).Value : Else : vmssMEDICO = "" : End If
                    If Not IsDBNull(Me.DGV17.Item(17, I).Value) Then : vmssN_MINSA = Me.DGV17.Item(17, I).Value : Else : vmssN_MINSA = "" : End If
                    If Not IsDBNull(Me.DGV17.Item(18, I).Value) Then : vmssID_E_MED = Me.DGV17.Item(18, I).Value : Else : vmssID_E_MED = 0 : End If
                    If Not IsDBNull(Me.DGV17.Item(19, I).Value) Then : vmssD_E_MED = Me.DGV17.Item(19, I).Value : Else : vmssD_E_MED = "" : End If
                    If Not IsDBNull(Me.DGV17.Item(20, I).Value) Then : vmssDIAGNOSTICO = Me.DGV17.Item(20, I).Value : Else : vmssDIAGNOSTICO = "" : End If
                    If Not IsDBNull(Me.DGV17.Item(21, I).Value) Then : vmssOBSERVACIONES = Me.DGV17.Item(21, I).Value : Else : vmssOBSERVACIONES = "" : End If
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub DGV17_DoubleClick(sender As Object, e As EventArgs) Handles DGV17.DoubleClick   'EDITAR
        Call VERIFICAR_ACCESOS("147") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmssID_SUBS = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Editar_Subsidios.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE SUBSIDIOS] WHERE ID_M_P = " & vID_M_P & "ORDER BY FECHA_I"
            Call CARGAR_DATAGRIDVIEW_17()
            Call ARMAR_DATAGRIDVIEW_17()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_17()
        End If
    End Sub
    Private Sub Button55_Click(sender As Object, e As EventArgs) Handles Button55.Click 'ELIMINAR
        Call VERIFICAR_ACCESOS("148") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmssID_SUBS = 0 Then
            MsgBox("Debe seleccionar el registro que desea Eliminar", vbInformation, "Mensaje del Sistema")
            Me.DGV16.Focus()
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Esta seguro de Eliminar el registro seleccionado?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Conexion.ABD_RRHH()
            Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
            Try
                MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE SUBSIDIOS] WHERE ID_SUBS = " & CInt(vmssID_SUBS) & ""
                MiSqlCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
            Call CARGAR_DATAGRIDVIEW_17()
            Call ARMAR_DATAGRIDVIEW_17()
        Else
            vmhiID_H = 0
        End If
    End Sub
    Private Sub Button58_Click(sender As Object, e As EventArgs) Handles Button58.Click
        Call VERIFICAR_ACCESOS("145") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Agregar_Subsidios_SS.ShowDialog()
    End Sub
    Private Sub Button56_Click(sender As Object, e As EventArgs) Handles Button56.Click 'EDITAR
        Call VERIFICAR_ACCESOS("147") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmssID_SUBS = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Editar_Subsidios.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE SUBSIDIOS] WHERE ID_M_P = " & vID_M_P & "ORDER BY FECHA_I"
            Call CARGAR_DATAGRIDVIEW_17()
            Call ARMAR_DATAGRIDVIEW_17()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_17()
        End If
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_17()
        Dim I As Integer
        For I = 0 To Me.DGV17.RowCount - 1
            If Me.DGV17.Item(0, I).Value = vmssID_SUBS Then
                Me.DGV17.Rows(I).Selected = True
                Me.DGV17.CurrentCell = Me.DGV17.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub Button57_Click(sender As Object, e As EventArgs) Handles Button57.Click 'NUEVO
        Call VERIFICAR_ACCESOS("146") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Agregar_Subsidios.ShowDialog()
    End Sub
    'VISTA MAESTRO DE INMUNIZACIONES
    Private Sub DGV19_Click(sender As Object, e As EventArgs) Handles DGV19.Click
        If Me.DGV19.RowCount > 0 Then
            For I = 0 To Me.DGV19.RowCount - 1
                If DGV19.Rows(I).Selected = True Then
                    If Not IsDBNull(Me.DGV19.Item(0, I).Value) Then : inID_M_INM = Me.DGV19.Item(0, I).Value : Else : inID_M_INM = 0 : End If
                    If Not IsDBNull(Me.DGV19.Item(1, I).Value) Then : inID_M_P = Me.DGV19.Item(1, I).Value : Else : inID_M_P = 0 : End If
                    If Not IsDBNull(Me.DGV19.Item(2, I).Value) Then : inID_VACUNA = Me.DGV19.Item(2, I).Value : Else : inID_VACUNA = 0 : End If
                    If Not IsDBNull(Me.DGV19.Item(3, I).Value) Then : inD_VACUNA = Me.DGV19.Item(3, I).Value : Else : inD_VACUNA = "" : End If
                    If Not IsDBNull(Me.DGV19.Item(4, I).Value) Then : inID_DOSIS = Me.DGV19.Item(4, I).Value : Else : inID_DOSIS = 0 : End If
                    If Not IsDBNull(Me.DGV19.Item(5, I).Value) Then : inD_DOSIS = Me.DGV19.Item(5, I).Value : Else : inD_DOSIS = "" : End If
                    If Not IsDBNull(Me.DGV19.Item(6, I).Value) Then : inFECHA_V = Me.DGV19.Item(6, I).Value : Else : inFECHA_V = "" : End If
                    Exit Sub
                End If
            Next
        End If
    End Sub
    'VISTA VISTA MAESTRO DE CONTRATOS Y EVALUACIONES
    Private Sub DGV18_Click(sender As Object, e As EventArgs) Handles DGV18.Click
        If Me.DGV18.RowCount > 0 Then
            For I = 0 To Me.DGV18.RowCount - 1
                If DGV18.Rows(I).Selected = True Then
                    If Not IsDBNull(Me.DGV18.Item(0, I).Value) Then : mcyeID_M_CYE = Me.DGV18.Item(0, I).Value : Else : mcyeID_M_CYE = 0 : End If
                    If Not IsDBNull(Me.DGV18.Item(1, I).Value) Then : mcyeID_M_P = Me.DGV18.Item(1, I).Value : Else : mcyeID_M_P = 0 : End If
                    If Not IsDBNull(Me.DGV18.Item(2, I).Value) Then : mcyeCODIGO = Me.DGV18.Item(2, I).Value : Else : mcyeCODIGO = "" : End If
                    If Not IsDBNull(Me.DGV18.Item(3, I).Value) Then : mcyeFEC_ING_PAME = Me.DGV18.Item(3, I).Value : Else : mcyeFEC_ING_PAME = "" : End If
                    If Not IsDBNull(Me.DGV18.Item(4, I).Value) Then : mcyeID_CONT = Me.DGV18.Item(4, I).Value : Else : mcyeID_CONT = 0 : End If
                    If Not IsDBNull(Me.DGV18.Item(5, I).Value) Then : mcyeDESCRIPCION = Me.DGV18.Item(5, I).Value : Else : mcyeDESCRIPCION = "" : End If
                    If Not IsDBNull(Me.DGV18.Item(6, I).Value) Then : mcyeFECHA_I = Me.DGV18.Item(6, I).Value : Else : mcyeFECHA_I = "" : End If
                    If Not IsDBNull(Me.DGV18.Item(7, I).Value) Then : mcyeFECHA_F = Me.DGV18.Item(7, I).Value : Else : mcyeFECHA_F = "" : End If
                    If Not IsDBNull(Me.DGV18.Item(8, I).Value) Then : mcyeFIRMADO = Me.DGV18.Item(8, I).Value : Else : mcyeFIRMADO = "" : End If
                    If Not IsDBNull(Me.DGV18.Item(9, I).Value) Then : mcyeFECHA_EVA = Me.DGV18.Item(9, I).Value : Else : mcyeFECHA_EVA = "" : End If
                    If Not IsDBNull(Me.DGV18.Item(10, I).Value) Then : mcyePROMEDIO_NOTA = Me.DGV18.Item(10, I).Value : Else : mcyePROMEDIO_NOTA = "" : End If
                    If Not IsDBNull(Me.DGV18.Item(11, I).Value) Then : mcyeID_C_E = Me.DGV18.Item(11, I).Value : Else : mcyeID_C_E = 0 : End If
                    If Not IsDBNull(Me.DGV18.Item(12, I).Value) Then : mcyeD_C_E = Me.DGV18.Item(12, I).Value : Else : mcyeD_C_E = "" : End If
                    If Not IsDBNull(Me.DGV18.Item(13, I).Value) Then : mcyeOBSERVACIONES = Me.DGV18.Item(13, I).Value : Else : mcyeOBSERVACIONES = "" : End If
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub DGV18_DoubleClick(sender As Object, e As EventArgs) Handles DGV18.DoubleClick   'EDITAR
        Call VERIFICAR_ACCESOS("151") : If HAY_ACCESO = False Then : Exit Sub : End If
        If mcyeID_M_CYE = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Editar_CyE.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [dbo].[VISTA MAESTRO DE CONTRATOS Y EVALUACIONES] WHERE ID_M_P = " & vID_M_P & " ORDER BY [CODIGO], [ID_CONT]"
            Call CARGAR_DATAGRIDVIEW_18()
            Call ARMAR_DATAGRIDVIEW_18()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_18()
        End If
    End Sub
    Private Sub Button59_Click(sender As Object, e As EventArgs) Handles Button59.Click 'ELIMINAR
        Call VERIFICAR_ACCESOS("152") : If HAY_ACCESO = False Then : Exit Sub : End If
        If mcyeID_M_CYE = 0 Then
            MsgBox("Debe seleccionar el registro que desea Eliminar", vbInformation, "Mensaje del Sistema")
            Me.DGV16.Focus()
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Esta seguro de Eliminar el registro seleccionado?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Conexion.ABD_RRHH()
            Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
            Try
                MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE CONTRATOS Y EVALUACIONES] WHERE ID_M_CYE = " & CInt(mcyeID_M_CYE) & ""
                MiSqlCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
            Call CARGAR_DATAGRIDVIEW_18()
            Call ARMAR_DATAGRIDVIEW_18()
        Else
            mcyeID_M_CYE = 0
        End If
    End Sub
    Private Sub Button62_Click(sender As Object, e As EventArgs) Handles Button62.Click
        Call VERIFICAR_ACCESOS("149") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Expedientes_Calcular_Fechas.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [dbo].[VISTA MAESTRO DE CONTRATOS Y EVALUACIONES] WHERE ID_M_P = " & vID_M_P & " ORDER BY [CODIGO], [ID_CONT]"
            Call CARGAR_DATAGRIDVIEW_18()
            Call ARMAR_DATAGRIDVIEW_18()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_18()
        End If
    End Sub
    Private Sub Button61_Click(sender As Object, e As EventArgs) Handles Button61.Click
        Call VERIFICAR_ACCESOS("150") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Agregar_CyE.ShowDialog()
    End Sub
    Private Sub Button60_Click(sender As Object, e As EventArgs) Handles Button60.Click 'EDITAR
        Call VERIFICAR_ACCESOS("151") : If HAY_ACCESO = False Then : Exit Sub : End If
        If mcyeID_M_CYE = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Editar_CyE.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [dbo].[VISTA MAESTRO DE CONTRATOS Y EVALUACIONES] WHERE ID_M_P = " & vID_M_P & " ORDER BY [CODIGO], [ID_CONT]"
            Call CARGAR_DATAGRIDVIEW_18()
            Call ARMAR_DATAGRIDVIEW_18()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_18()
        End If
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_18()
        Dim I As Integer
        For I = 0 To Me.DGV18.RowCount - 1
            If Me.DGV18.Item(0, I).Value = mcyeID_M_CYE Then
                Me.DGV18.Rows(I).Selected = True
                Me.DGV18.CurrentCell = Me.DGV18.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub DGV11_RowPrePaint(sender As Object, e As DataGridViewRowPrePaintEventArgs) Handles DGV11.RowPrePaint
        Dim style As New DataGridViewCellStyle
        style.Font = New Font(DGV11.Font, FontStyle.Bold)
        For Each Row As DataGridViewRow In DGV11.Rows
            If IsDBNull(Row.Cells(11).Value) Then

            Else
                If Row.Cells(11).Value = False Then
                    Row.DefaultCellStyle.ForeColor = Color.Red
                End If
            End If
        Next
    End Sub
    Private Sub ComboBox33_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox33.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox26_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox26.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox29_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox29.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub Button45_Click(sender As Object, e As EventArgs) Handles Button45.Click
        Call VERIFICAR_ACCESOS("103") : If HAY_ACCESO = False Then : Exit Sub : End If
        If Me.TextBox35.Text = "" Or Me.TextBox35.Text = "0" Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Me.Button1.Focus()
        Else
            Call ACTIVAR_03()
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim tp As TabPage = TabGeneral.TabPages(2)
        TabGeneral.SelectedTab = tp
        Me.Label4.Text = Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Text
        Call NO_ACTIVAR()
        Call SELECCIONAR_CADENA()
        TabGeneral.Focus()
        Call DESACTIVAR_TODO()
    End Sub
    Private Sub Button47_Click_1(sender As Object, e As EventArgs) Handles Button47.Click
        Frm_Expedientes_Activos_Jefe_Inmediato.ShowDialog()
        If CERRAR = False Then
            Me.TextBox44.Text = bEMPLEADO
        End If
    End Sub
    Private Sub TextBox43_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox43.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub DateTimePicker8_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker8.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub DateTimePicker9_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker9.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub Button63_Click(sender As Object, e As EventArgs) Handles Button63.Click
        Frm_Expedientes_Activos_Prefijos.ShowDialog()
        If CERRAR = False Then
            If bPREFIJO <> "" Then
                Me.TextBox28.Text = bPREFIJO
            End If
        End If
    End Sub
    Private Sub Button66_Click(sender As Object, e As EventArgs) Handles Button66.Click
        Call VERIFICAR_ACCESOS("315") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Agregar_Inm.ShowDialog()
    End Sub
    Private Sub Button65_Click(sender As Object, e As EventArgs) Handles Button65.Click
        Call VERIFICAR_ACCESOS("316") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Editar_Inm.ShowDialog()
    End Sub
    Private Sub Button64_Click(sender As Object, e As EventArgs) Handles Button64.Click
        Call VERIFICAR_ACCESOS("317") : If HAY_ACCESO = False Then : Exit Sub : End If
        If inID_M_INM = 0 Then
            MsgBox("Debe seleccionar el registro que desea Eliminar", vbInformation, "Mensaje del Sistema")
            Me.DGV9.Focus()
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Esta seguro de Eliminar el registro seleccionado?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Conexion.ABD_RRHH()
            Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
            Try
                MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE INMUNIZACIONES] WHERE ID_M_INM = " & CInt(inID_M_INM) & ""
                MiSqlCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
            Call CARGAR_DATAGRIDVIEW_19()
            Call ARMAR_DATAGRIDVIEW_19()
        Else
            inID_M_INM = 0
        End If
    End Sub
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call VERIFICAR_ACCESOS("075") : If HAY_ACCESO = False Then : Exit Sub : End If
            cEMPLEADO = CInt(Val(TextBox2.Text))
            TextBox2.Text = Format(CInt(cEMPLEADO), "0000000000")
            If Len(Me.TextBox2.Text) = 0 Or (Me.TextBox2.Text = "0000000000") Then
                MsgBox("El código que se digitó no es válido", vbInformation, "Mensaje del Sistema")
                Me.TextBox2.Focus()
                Exit Sub
            Else
                Call BUSCAR_DATO_1()
                If DATO_1 = True Then
                    Dim tp As TabPage = TabGeneral.TabPages(1)
                    tp = TabGeneral.TabPages(1)
                    TabGeneral.SelectedTab = tp
                    Me.Label4.Text = Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Text
                    Call NO_ACTIVAR()
                    Call SELECCIONAR_CADENA()
                    Call DESACTIVAR_TODO()
                    TabGeneral.Focus()
                    '------------------------------------------------------------------------------------
                    Call BUSCAR_UBICACION()
                    '------------------------------------------------------------------------------------
                    Me.TextBox2.SelectAll()
                    Me.TextBox2.Focus()
                Else
                    MsgBox("El código digitado no se encuentra", vbInformation, "Mensaje del Sistema")
                    Me.TextBox2.SelectAll()
                    Me.TextBox2.Focus()
                End If
            End If
        End If
    End Sub
    Dim DATO_1 As Boolean
    Private Sub BUSCAR_DATO_1()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            If isuID_TIPO_USUARIO = 5 Then
                CADENAsql = "SELECT ID_M_P, CODIGO, ID_ESTADO_P FROM [dbo].[VISTA MAESTRO DE PERSONAS] WHERE CODIGO = '" & Me.TextBox2.Text & "' AND ID_ESTADO_P = 1 AND D_ADMIN = '" & isuUSUARIO & "'"
            Else
                CADENAsql = "SELECT ID_M_P, CODIGO, ID_ESTADO_P FROM [dbo].[VISTA MAESTRO DE PERSONAS] WHERE CODIGO = '" & Me.TextBox2.Text & "' AND ID_ESTADO_P = 1"
            End If
            MiDataAdapter = New SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                vID_M_P = MiDataTable.Rows(0).Item(0).ToString
                DATO_1 = True
            Else
                vID_M_P = 0
                DATO_1 = False
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub DGV19_DoubleClick(sender As Object, e As EventArgs) Handles DGV19.DoubleClick
        Call VERIFICAR_ACCESOS("316") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Editar_Inm.ShowDialog()
    End Sub
End Class