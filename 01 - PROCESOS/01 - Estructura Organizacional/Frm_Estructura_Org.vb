Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Estructura_Org
    Dim CONTROL_SELECCION_DGV As Integer
    Private Sub Frm_Estructura_Org_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            ESTRUCTURA = False
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Estructura_Org_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Frm_Principal.Timer1.Stop()
        ESTRUCTURA = True
        RemoveHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
        RemoveHandler Me.ComboBox2.SelectedIndexChanged, AddressOf ComboBox2_SelectedIndexChanged
        RemoveHandler Me.ComboBox3.SelectedIndexChanged, AddressOf ComboBox3_SelectedIndexChanged
        RemoveHandler Me.ComboBox4.SelectedIndexChanged, AddressOf ComboBox4_SelectedIndexChanged
        RemoveHandler Me.ComboBox5.SelectedIndexChanged, AddressOf ComboBox5_SelectedIndexChanged
        RemoveHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
        RemoveHandler Me.ComboBox7.SelectedIndexChanged, AddressOf ComboBox7_SelectedIndexChanged
        RemoveHandler Me.ComboBox8.SelectedIndexChanged, AddressOf ComboBox8_SelectedIndexChanged
        RemoveHandler CheckBox1.CheckedChanged, AddressOf CheckBox1_CheckedChanged
        RemoveHandler CheckBox2.CheckedChanged, AddressOf CheckBox2_CheckedChanged
        RemoveHandler CheckBox3.CheckedChanged, AddressOf CheckBox3_CheckedChanged
        RemoveHandler CheckBox4.CheckedChanged, AddressOf CheckBox4_CheckedChanged
        RemoveHandler CheckBox5.CheckedChanged, AddressOf CheckBox5_CheckedChanged
        RemoveHandler CheckBox6.CheckedChanged, AddressOf CheckBox6_CheckedChanged
        RemoveHandler CheckBox7.CheckedChanged, AddressOf CheckBox7_CheckedChanged
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_1()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_2()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_3()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_4()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_5()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_6()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_7()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_8()
        AddHandler Me.CheckBox7.CheckedChanged, AddressOf CheckBox7_CheckedChanged
        AddHandler Me.CheckBox6.CheckedChanged, AddressOf CheckBox6_CheckedChanged
        AddHandler Me.CheckBox5.CheckedChanged, AddressOf CheckBox5_CheckedChanged
        AddHandler Me.CheckBox4.CheckedChanged, AddressOf CheckBox4_CheckedChanged
        AddHandler Me.CheckBox3.CheckedChanged, AddressOf CheckBox3_CheckedChanged
        AddHandler Me.CheckBox2.CheckedChanged, AddressOf CheckBox2_CheckedChanged
        AddHandler Me.CheckBox1.CheckedChanged, AddressOf CheckBox1_CheckedChanged
        AddHandler Me.ComboBox8.SelectedIndexChanged, AddressOf ComboBox8_SelectedIndexChanged
        AddHandler Me.ComboBox7.SelectedIndexChanged, AddressOf ComboBox7_SelectedIndexChanged
        AddHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
        AddHandler Me.ComboBox5.SelectedIndexChanged, AddressOf ComboBox5_SelectedIndexChanged
        AddHandler Me.ComboBox4.SelectedIndexChanged, AddressOf ComboBox4_SelectedIndexChanged
        AddHandler Me.ComboBox3.SelectedIndexChanged, AddressOf ComboBox3_SelectedIndexChanged
        AddHandler Me.ComboBox2.SelectedIndexChanged, AddressOf ComboBox2_SelectedIndexChanged
        AddHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
        Call ARMAR_CADENA_NIVELES()
        If isuID_TIPO_USUARIO = 5 Then
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        Else
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        End If
        Call CARGAR_DATAGRIDVIEW_3()
        Call ARMAR_DATAGRIDVIEW_3()
    End Sub
    Private Sub Button7_Click(sender As Object, e As System.EventArgs) Handles Button7.Click
        ESTRUCTURA = False
        Frm_Principal.Timer1.Start()
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_8()
        If Me.ComboBox7.Items.Count > 0 Then
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Conexion.ABD_RRHH()
            Dim MiDataTable As New DataTable
            Dim MiSqlDataAdapter As SqlDataAdapter
            Try
                MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTRUCTURA NIVEL 8] where ID_N7 = " & Me.ComboBox7.SelectedValue & " order by ORDEN", Conexion.CxRRHH)
                MiDataTable.Clear()
                MiSqlDataAdapter.Fill(MiDataTable)
                If MiDataTable.Rows.Count <> 0 Then
                    Me.ComboBox8.Font = New Font("Microsoft Sans Serif", 8.25)
                    Me.ComboBox8.ForeColor = Color.Black
                    Me.CheckBox7.Checked = True : Me.CheckBox7.Enabled = True
                    With Me.ComboBox8
                        .DataSource = MiDataTable
                        .DisplayMember = "DESCRIPCION"
                        .ValueMember = "ID_N8"
                    End With
                Else
                    Me.ComboBox8.DataSource = Nothing
                    Me.ComboBox8.ForeColor = Color.Red
                    Me.ComboBox8.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
                    Me.ComboBox8.Text = "-- SIN DATOS -- "
                    Me.CheckBox7.Checked = False : Me.CheckBox7.Enabled = False
                End If
            Catch ex As Exception
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
        Else
            Me.ComboBox8.DataSource = Nothing
            Me.ComboBox8.ForeColor = Color.Red
            Me.ComboBox8.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
            Me.ComboBox8.Text = "-- SIN DATOS -- "
            Me.CheckBox7.Checked = False : Me.CheckBox7.Enabled = False
        End If
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_7()
        If Me.ComboBox6.Items.Count > 0 Then
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Conexion.ABD_RRHH()
            Dim MiDataTable As New DataTable
            Dim MiSqlDataAdapter As SqlDataAdapter
            Try
                MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTRUCTURA NIVEL 7] where ID_N6 = " & Me.ComboBox6.SelectedValue & " order by ORDEN", Conexion.CxRRHH)
                MiDataTable.Clear()
                MiSqlDataAdapter.Fill(MiDataTable)
                If MiDataTable.Rows.Count <> 0 Then
                    Me.ComboBox7.Font = New Font("Microsoft Sans Serif", 8.25)
                    Me.ComboBox7.ForeColor = Color.Black
                    Me.CheckBox6.Checked = True : Me.CheckBox6.Enabled = True
                    With Me.ComboBox7
                        .DataSource = MiDataTable
                        .DisplayMember = "DESCRIPCION"
                        .ValueMember = "ID_N7"
                    End With
                Else
                    Me.ComboBox7.DataSource = Nothing
                    Me.ComboBox7.ForeColor = Color.Red
                    Me.ComboBox7.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
                    Me.ComboBox7.Text = "-- SIN DATOS -- "
                    Me.CheckBox6.Checked = False : Me.CheckBox6.Enabled = False
                End If
            Catch ex As Exception
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
        Else
            Me.ComboBox7.DataSource = Nothing
            Me.ComboBox7.ForeColor = Color.Red
            Me.ComboBox7.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
            Me.ComboBox7.Text = "-- SIN DATOS -- "
            Me.CheckBox6.Checked = False : Me.CheckBox6.Enabled = False
        End If
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_6()
        If Me.ComboBox5.Items.Count > 0 Then
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Conexion.ABD_RRHH()
            Dim MiDataTable As New DataTable
            Dim MiSqlDataAdapter As SqlDataAdapter
            Try
                MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTRUCTURA NIVEL 6] where ID_N5 = " & Me.ComboBox5.SelectedValue & " order by ORDEN", Conexion.CxRRHH)
                MiDataTable.Clear()
                MiSqlDataAdapter.Fill(MiDataTable)
                If MiDataTable.Rows.Count <> 0 Then
                    Me.ComboBox6.Font = New Font("Microsoft Sans Serif", 8.25)
                    Me.ComboBox6.ForeColor = Color.Black
                    Me.CheckBox5.Checked = True : Me.CheckBox5.Enabled = True
                    With Me.ComboBox6
                        .DataSource = MiDataTable
                        .DisplayMember = "DESCRIPCION"
                        .ValueMember = "ID_N6"
                    End With
                Else
                    Me.ComboBox6.DataSource = Nothing
                    Me.ComboBox6.ForeColor = Color.Red
                    Me.ComboBox6.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
                    Me.ComboBox6.Text = "-- SIN DATOS -- "
                    Me.CheckBox5.Checked = False : Me.CheckBox5.Enabled = False
                End If
            Catch ex As Exception
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
        Else
            Me.ComboBox6.DataSource = Nothing
            Me.ComboBox6.ForeColor = Color.Red
            Me.ComboBox6.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
            Me.ComboBox6.Text = "-- SIN DATOS -- "
            Me.CheckBox5.Checked = False : Me.CheckBox5.Enabled = False
        End If
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_5()
        If Me.ComboBox4.Items.Count > 0 Then
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Conexion.ABD_RRHH()
            Dim MiDataTable As New DataTable
            Dim MiSqlDataAdapter As SqlDataAdapter
            Try
                MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTRUCTURA NIVEL 5] where ID_N4 = " & Me.ComboBox4.SelectedValue & " order by ORDEN", Conexion.CxRRHH)
                MiDataTable.Clear()
                MiSqlDataAdapter.Fill(MiDataTable)
                If MiDataTable.Rows.Count <> 0 Then
                    Me.ComboBox5.Font = New Font("Microsoft Sans Serif", 8.25)
                    Me.ComboBox5.ForeColor = Color.Black
                    Me.CheckBox4.Checked = True : Me.CheckBox4.Enabled = True
                    With Me.ComboBox5
                        .DataSource = MiDataTable
                        .DisplayMember = "DESCRIPCION"
                        .ValueMember = "ID_N5"
                    End With
                Else
                    Me.ComboBox5.DataSource = Nothing
                    Me.ComboBox5.ForeColor = Color.Red
                    Me.ComboBox5.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
                    Me.ComboBox5.Text = "-- SIN DATOS -- "
                    Me.CheckBox4.Checked = False : Me.CheckBox4.Enabled = False
                End If
            Catch ex As Exception
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
        Else
            Me.ComboBox5.DataSource = Nothing
            Me.ComboBox5.ForeColor = Color.Red
            Me.ComboBox5.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
            Me.ComboBox5.Text = "-- SIN DATOS -- "
            Me.CheckBox4.Checked = False : Me.CheckBox4.Enabled = False
        End If
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_4()
        If Me.ComboBox3.Items.Count > 0 Then
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Conexion.ABD_RRHH()
            Dim MiDataTable As New DataTable
            Dim MiSqlDataAdapter As SqlDataAdapter
            Try
                MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTRUCTURA NIVEL 4] where ID_N3 = " & Me.ComboBox3.SelectedValue & " order by ORDEN", Conexion.CxRRHH)
                MiDataTable.Clear()
                MiSqlDataAdapter.Fill(MiDataTable)
                If MiDataTable.Rows.Count <> 0 Then
                    Me.ComboBox4.Font = New Font("Microsoft Sans Serif", 8.25)
                    Me.ComboBox4.ForeColor = Color.Black
                    Me.CheckBox3.Checked = True : Me.CheckBox3.Enabled = True
                    With Me.ComboBox4
                        .DataSource = MiDataTable
                        .DisplayMember = "DESCRIPCION"
                        .ValueMember = "ID_N4"
                    End With
                Else
                    Me.ComboBox4.DataSource = Nothing
                    Me.ComboBox4.ForeColor = Color.Red
                    Me.ComboBox4.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
                    Me.ComboBox4.Text = "-- SIN DATOS -- "
                    Me.CheckBox3.Checked = False : Me.CheckBox3.Enabled = False
                End If
            Catch ex As Exception
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
        Else
            Me.ComboBox4.DataSource = Nothing
            Me.ComboBox4.ForeColor = Color.Red
            Me.ComboBox4.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
            Me.ComboBox4.Text = "-- SIN DATOS -- "
            Me.CheckBox3.Checked = False : Me.CheckBox3.Enabled = False
        End If
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_3()
        If Me.ComboBox2.Items.Count > 0 Then
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Conexion.ABD_RRHH()
            Dim MiDataTable As New DataTable
            Dim MiSqlDataAdapter As SqlDataAdapter
            Try
                MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTRUCTURA NIVEL 3] where ID_N2 = " & Me.ComboBox2.SelectedValue & " order by ORDEN", Conexion.CxRRHH)
                MiDataTable.Clear()
                MiSqlDataAdapter.Fill(MiDataTable)
                If MiDataTable.Rows.Count <> 0 Then
                    Me.ComboBox3.Font = New Font("Microsoft Sans Serif", 8.25)
                    Me.ComboBox3.ForeColor = Color.Black
                    Me.CheckBox2.Checked = True : Me.CheckBox2.Enabled = True
                    With Me.ComboBox3
                        .DataSource = MiDataTable
                        .DisplayMember = "DESCRIPCION"
                        .ValueMember = "ID_N3"
                    End With
                Else
                    Me.ComboBox3.DataSource = Nothing
                    Me.ComboBox3.ForeColor = Color.Red
                    Me.ComboBox3.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
                    Me.ComboBox3.Text = "-- SIN DATOS -- "
                    Me.CheckBox2.Checked = False : Me.CheckBox2.Enabled = False
                End If
            Catch ex As Exception
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
        Else
            Me.ComboBox3.DataSource = Nothing
            Me.ComboBox3.ForeColor = Color.Red
            Me.ComboBox3.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
            Me.ComboBox3.Text = "-- SIN DATOS -- "
            Me.CheckBox2.Checked = False : Me.CheckBox2.Enabled = False
        End If
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_2()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTRUCTURA NIVEL 2] where ID_N1 = " & Me.ComboBox1.SelectedValue & " order by ORDEN", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count <> 0 Then
                Me.ComboBox2.Font = New Font("Microsoft Sans Serif", 8.25)
                Me.ComboBox2.ForeColor = Color.Black
                Me.CheckBox1.Checked = True : Me.CheckBox1.Enabled = True
                With Me.ComboBox2
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_N2"
                End With
            Else
                Me.ComboBox2.DataSource = Nothing
                Me.ComboBox2.ForeColor = Color.Red
                Me.ComboBox2.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
                Me.ComboBox2.Text = "-- SIN DATOS --"
                Me.CheckBox1.Checked = False : Me.CheckBox1.Enabled = False
            End If
        Catch ex As Exception
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
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTRUCTURA NIVEL 1] order by ORDEN", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_N1"
            End With

        Catch ex As Exception
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        RemoveHandler Me.ComboBox2.SelectedIndexChanged, AddressOf ComboBox2_SelectedIndexChanged
        RemoveHandler Me.ComboBox3.SelectedIndexChanged, AddressOf ComboBox3_SelectedIndexChanged
        RemoveHandler Me.ComboBox4.SelectedIndexChanged, AddressOf ComboBox4_SelectedIndexChanged
        RemoveHandler Me.ComboBox5.SelectedIndexChanged, AddressOf ComboBox5_SelectedIndexChanged
        RemoveHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
        RemoveHandler Me.ComboBox7.SelectedIndexChanged, AddressOf ComboBox7_SelectedIndexChanged
        RemoveHandler Me.ComboBox8.SelectedIndexChanged, AddressOf ComboBox8_SelectedIndexChanged
        RemoveHandler CheckBox1.CheckedChanged, AddressOf CheckBox1_CheckedChanged
        RemoveHandler CheckBox2.CheckedChanged, AddressOf CheckBox2_CheckedChanged
        RemoveHandler CheckBox3.CheckedChanged, AddressOf CheckBox3_CheckedChanged
        RemoveHandler CheckBox4.CheckedChanged, AddressOf CheckBox4_CheckedChanged
        RemoveHandler CheckBox5.CheckedChanged, AddressOf CheckBox5_CheckedChanged
        RemoveHandler CheckBox6.CheckedChanged, AddressOf CheckBox6_CheckedChanged
        RemoveHandler CheckBox7.CheckedChanged, AddressOf CheckBox7_CheckedChanged
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_2()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_3()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_4()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_5()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_6()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_7()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_8()
        AddHandler Me.CheckBox7.CheckedChanged, AddressOf CheckBox7_CheckedChanged
        AddHandler Me.CheckBox6.CheckedChanged, AddressOf CheckBox6_CheckedChanged
        AddHandler Me.CheckBox5.CheckedChanged, AddressOf CheckBox5_CheckedChanged
        AddHandler Me.CheckBox4.CheckedChanged, AddressOf CheckBox4_CheckedChanged
        AddHandler Me.CheckBox3.CheckedChanged, AddressOf CheckBox3_CheckedChanged
        AddHandler Me.CheckBox2.CheckedChanged, AddressOf CheckBox2_CheckedChanged
        AddHandler Me.CheckBox1.CheckedChanged, AddressOf CheckBox1_CheckedChanged
        AddHandler Me.ComboBox8.SelectedIndexChanged, AddressOf ComboBox8_SelectedIndexChanged
        AddHandler Me.ComboBox7.SelectedIndexChanged, AddressOf ComboBox7_SelectedIndexChanged
        AddHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
        AddHandler Me.ComboBox5.SelectedIndexChanged, AddressOf ComboBox5_SelectedIndexChanged
        AddHandler Me.ComboBox4.SelectedIndexChanged, AddressOf ComboBox4_SelectedIndexChanged
        AddHandler Me.ComboBox3.SelectedIndexChanged, AddressOf ComboBox3_SelectedIndexChanged
        AddHandler Me.ComboBox2.SelectedIndexChanged, AddressOf ComboBox2_SelectedIndexChanged
        Call ARMAR_CADENA_NIVELES()
        If isuID_TIPO_USUARIO = 5 Then
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        Else
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        End If
        Call CARGAR_DATAGRIDVIEW_3()
        Call ARMAR_DATAGRIDVIEW_3()
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        RemoveHandler Me.ComboBox3.SelectedIndexChanged, AddressOf ComboBox3_SelectedIndexChanged
        RemoveHandler Me.ComboBox4.SelectedIndexChanged, AddressOf ComboBox4_SelectedIndexChanged
        RemoveHandler Me.ComboBox5.SelectedIndexChanged, AddressOf ComboBox5_SelectedIndexChanged
        RemoveHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
        RemoveHandler Me.ComboBox7.SelectedIndexChanged, AddressOf ComboBox7_SelectedIndexChanged
        RemoveHandler Me.ComboBox8.SelectedIndexChanged, AddressOf ComboBox8_SelectedIndexChanged
        RemoveHandler CheckBox1.CheckedChanged, AddressOf CheckBox1_CheckedChanged
        RemoveHandler CheckBox2.CheckedChanged, AddressOf CheckBox2_CheckedChanged
        RemoveHandler CheckBox3.CheckedChanged, AddressOf CheckBox3_CheckedChanged
        RemoveHandler CheckBox4.CheckedChanged, AddressOf CheckBox4_CheckedChanged
        RemoveHandler CheckBox5.CheckedChanged, AddressOf CheckBox5_CheckedChanged
        RemoveHandler CheckBox6.CheckedChanged, AddressOf CheckBox6_CheckedChanged
        RemoveHandler CheckBox7.CheckedChanged, AddressOf CheckBox7_CheckedChanged
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_3()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_4()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_5()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_6()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_7()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_8()
        AddHandler Me.CheckBox7.CheckedChanged, AddressOf CheckBox7_CheckedChanged
        AddHandler Me.CheckBox6.CheckedChanged, AddressOf CheckBox6_CheckedChanged
        AddHandler Me.CheckBox5.CheckedChanged, AddressOf CheckBox5_CheckedChanged
        AddHandler Me.CheckBox4.CheckedChanged, AddressOf CheckBox4_CheckedChanged
        AddHandler Me.CheckBox3.CheckedChanged, AddressOf CheckBox3_CheckedChanged
        AddHandler Me.CheckBox2.CheckedChanged, AddressOf CheckBox2_CheckedChanged
        AddHandler Me.CheckBox1.CheckedChanged, AddressOf CheckBox1_CheckedChanged
        AddHandler Me.ComboBox8.SelectedIndexChanged, AddressOf ComboBox8_SelectedIndexChanged
        AddHandler Me.ComboBox7.SelectedIndexChanged, AddressOf ComboBox7_SelectedIndexChanged
        AddHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
        AddHandler Me.ComboBox5.SelectedIndexChanged, AddressOf ComboBox5_SelectedIndexChanged
        AddHandler Me.ComboBox4.SelectedIndexChanged, AddressOf ComboBox4_SelectedIndexChanged
        AddHandler Me.ComboBox3.SelectedIndexChanged, AddressOf ComboBox3_SelectedIndexChanged
        Call ARMAR_CADENA_NIVELES()
        If isuID_TIPO_USUARIO = 5 Then
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        Else
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        End If
        Call CARGAR_DATAGRIDVIEW_3()
        Call ARMAR_DATAGRIDVIEW_3()
    End Sub
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        RemoveHandler Me.ComboBox4.SelectedIndexChanged, AddressOf ComboBox4_SelectedIndexChanged
        RemoveHandler Me.ComboBox5.SelectedIndexChanged, AddressOf ComboBox5_SelectedIndexChanged
        RemoveHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
        RemoveHandler Me.ComboBox7.SelectedIndexChanged, AddressOf ComboBox7_SelectedIndexChanged
        RemoveHandler Me.ComboBox8.SelectedIndexChanged, AddressOf ComboBox8_SelectedIndexChanged
        RemoveHandler CheckBox1.CheckedChanged, AddressOf CheckBox1_CheckedChanged
        RemoveHandler CheckBox2.CheckedChanged, AddressOf CheckBox2_CheckedChanged
        RemoveHandler CheckBox3.CheckedChanged, AddressOf CheckBox3_CheckedChanged
        RemoveHandler CheckBox4.CheckedChanged, AddressOf CheckBox4_CheckedChanged
        RemoveHandler CheckBox5.CheckedChanged, AddressOf CheckBox5_CheckedChanged
        RemoveHandler CheckBox6.CheckedChanged, AddressOf CheckBox6_CheckedChanged
        RemoveHandler CheckBox7.CheckedChanged, AddressOf CheckBox7_CheckedChanged
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_4()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_5()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_6()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_7()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_8()
        AddHandler Me.CheckBox7.CheckedChanged, AddressOf CheckBox7_CheckedChanged
        AddHandler Me.CheckBox6.CheckedChanged, AddressOf CheckBox6_CheckedChanged
        AddHandler Me.CheckBox5.CheckedChanged, AddressOf CheckBox5_CheckedChanged
        AddHandler Me.CheckBox4.CheckedChanged, AddressOf CheckBox4_CheckedChanged
        AddHandler Me.CheckBox3.CheckedChanged, AddressOf CheckBox3_CheckedChanged
        AddHandler Me.CheckBox2.CheckedChanged, AddressOf CheckBox2_CheckedChanged
        AddHandler Me.CheckBox1.CheckedChanged, AddressOf CheckBox1_CheckedChanged
        AddHandler Me.ComboBox8.SelectedIndexChanged, AddressOf ComboBox8_SelectedIndexChanged
        AddHandler Me.ComboBox7.SelectedIndexChanged, AddressOf ComboBox7_SelectedIndexChanged
        AddHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
        AddHandler Me.ComboBox5.SelectedIndexChanged, AddressOf ComboBox5_SelectedIndexChanged
        AddHandler Me.ComboBox4.SelectedIndexChanged, AddressOf ComboBox4_SelectedIndexChanged
        Call ARMAR_CADENA_NIVELES()
        If isuID_TIPO_USUARIO = 5 Then
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        Else
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        End If
        Call CARGAR_DATAGRIDVIEW_3()
        Call ARMAR_DATAGRIDVIEW_3()
    End Sub
    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        RemoveHandler Me.ComboBox5.SelectedIndexChanged, AddressOf ComboBox5_SelectedIndexChanged
        RemoveHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
        RemoveHandler Me.ComboBox7.SelectedIndexChanged, AddressOf ComboBox7_SelectedIndexChanged
        RemoveHandler Me.ComboBox8.SelectedIndexChanged, AddressOf ComboBox8_SelectedIndexChanged
        RemoveHandler CheckBox1.CheckedChanged, AddressOf CheckBox1_CheckedChanged
        RemoveHandler CheckBox2.CheckedChanged, AddressOf CheckBox2_CheckedChanged
        RemoveHandler CheckBox3.CheckedChanged, AddressOf CheckBox3_CheckedChanged
        RemoveHandler CheckBox4.CheckedChanged, AddressOf CheckBox4_CheckedChanged
        RemoveHandler CheckBox5.CheckedChanged, AddressOf CheckBox5_CheckedChanged
        RemoveHandler CheckBox6.CheckedChanged, AddressOf CheckBox6_CheckedChanged
        RemoveHandler CheckBox7.CheckedChanged, AddressOf CheckBox7_CheckedChanged
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_5()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_6()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_7()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_8()
        AddHandler Me.CheckBox7.CheckedChanged, AddressOf CheckBox7_CheckedChanged
        AddHandler Me.CheckBox6.CheckedChanged, AddressOf CheckBox6_CheckedChanged
        AddHandler Me.CheckBox5.CheckedChanged, AddressOf CheckBox5_CheckedChanged
        AddHandler Me.CheckBox4.CheckedChanged, AddressOf CheckBox4_CheckedChanged
        AddHandler Me.CheckBox3.CheckedChanged, AddressOf CheckBox3_CheckedChanged
        AddHandler Me.CheckBox2.CheckedChanged, AddressOf CheckBox2_CheckedChanged
        AddHandler Me.CheckBox1.CheckedChanged, AddressOf CheckBox1_CheckedChanged
        AddHandler Me.ComboBox8.SelectedIndexChanged, AddressOf ComboBox8_SelectedIndexChanged
        AddHandler Me.ComboBox7.SelectedIndexChanged, AddressOf ComboBox7_SelectedIndexChanged
        AddHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
        AddHandler Me.ComboBox5.SelectedIndexChanged, AddressOf ComboBox5_SelectedIndexChanged
        Call ARMAR_CADENA_NIVELES()
        If isuID_TIPO_USUARIO = 5 Then
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        Else
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        End If
        Call CARGAR_DATAGRIDVIEW_3()
        Call ARMAR_DATAGRIDVIEW_3()
    End Sub
    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        RemoveHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
        RemoveHandler Me.ComboBox7.SelectedIndexChanged, AddressOf ComboBox7_SelectedIndexChanged
        RemoveHandler Me.ComboBox8.SelectedIndexChanged, AddressOf ComboBox8_SelectedIndexChanged
        RemoveHandler CheckBox1.CheckedChanged, AddressOf CheckBox1_CheckedChanged
        RemoveHandler CheckBox2.CheckedChanged, AddressOf CheckBox2_CheckedChanged
        RemoveHandler CheckBox3.CheckedChanged, AddressOf CheckBox3_CheckedChanged
        RemoveHandler CheckBox4.CheckedChanged, AddressOf CheckBox4_CheckedChanged
        RemoveHandler CheckBox5.CheckedChanged, AddressOf CheckBox5_CheckedChanged
        RemoveHandler CheckBox6.CheckedChanged, AddressOf CheckBox6_CheckedChanged
        RemoveHandler CheckBox7.CheckedChanged, AddressOf CheckBox7_CheckedChanged
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_6()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_7()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_8()
        AddHandler Me.CheckBox7.CheckedChanged, AddressOf CheckBox7_CheckedChanged
        AddHandler Me.CheckBox6.CheckedChanged, AddressOf CheckBox6_CheckedChanged
        AddHandler Me.CheckBox5.CheckedChanged, AddressOf CheckBox5_CheckedChanged
        AddHandler Me.CheckBox4.CheckedChanged, AddressOf CheckBox4_CheckedChanged
        AddHandler Me.CheckBox3.CheckedChanged, AddressOf CheckBox3_CheckedChanged
        AddHandler Me.CheckBox2.CheckedChanged, AddressOf CheckBox2_CheckedChanged
        AddHandler Me.CheckBox1.CheckedChanged, AddressOf CheckBox1_CheckedChanged
        AddHandler Me.ComboBox8.SelectedIndexChanged, AddressOf ComboBox8_SelectedIndexChanged
        AddHandler Me.ComboBox7.SelectedIndexChanged, AddressOf ComboBox7_SelectedIndexChanged
        AddHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
        Call ARMAR_CADENA_NIVELES()
        If isuID_TIPO_USUARIO = 5 Then
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        Else
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        End If
        Call CARGAR_DATAGRIDVIEW_3()
        Call ARMAR_DATAGRIDVIEW_3()
    End Sub
    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        RemoveHandler Me.ComboBox7.SelectedIndexChanged, AddressOf ComboBox7_SelectedIndexChanged
        RemoveHandler Me.ComboBox8.SelectedIndexChanged, AddressOf ComboBox8_SelectedIndexChanged
        RemoveHandler CheckBox1.CheckedChanged, AddressOf CheckBox1_CheckedChanged
        RemoveHandler CheckBox2.CheckedChanged, AddressOf CheckBox2_CheckedChanged
        RemoveHandler CheckBox3.CheckedChanged, AddressOf CheckBox3_CheckedChanged
        RemoveHandler CheckBox4.CheckedChanged, AddressOf CheckBox4_CheckedChanged
        RemoveHandler CheckBox5.CheckedChanged, AddressOf CheckBox5_CheckedChanged
        RemoveHandler CheckBox6.CheckedChanged, AddressOf CheckBox6_CheckedChanged
        RemoveHandler CheckBox7.CheckedChanged, AddressOf CheckBox7_CheckedChanged
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_7()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_8()
        AddHandler Me.CheckBox7.CheckedChanged, AddressOf CheckBox7_CheckedChanged
        AddHandler Me.CheckBox6.CheckedChanged, AddressOf CheckBox6_CheckedChanged
        AddHandler Me.CheckBox5.CheckedChanged, AddressOf CheckBox5_CheckedChanged
        AddHandler Me.CheckBox4.CheckedChanged, AddressOf CheckBox4_CheckedChanged
        AddHandler Me.CheckBox3.CheckedChanged, AddressOf CheckBox3_CheckedChanged
        AddHandler Me.CheckBox2.CheckedChanged, AddressOf CheckBox2_CheckedChanged
        AddHandler Me.CheckBox1.CheckedChanged, AddressOf CheckBox1_CheckedChanged
        AddHandler Me.ComboBox8.SelectedIndexChanged, AddressOf ComboBox8_SelectedIndexChanged
        AddHandler Me.ComboBox7.SelectedIndexChanged, AddressOf ComboBox7_SelectedIndexChanged
        Call ARMAR_CADENA_NIVELES()
        If isuID_TIPO_USUARIO = 5 Then
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        Else
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        End If
        Call CARGAR_DATAGRIDVIEW_3()
        Call ARMAR_DATAGRIDVIEW_3()
    End Sub
    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        RemoveHandler Me.ComboBox8.SelectedIndexChanged, AddressOf ComboBox8_SelectedIndexChanged
        RemoveHandler CheckBox1.CheckedChanged, AddressOf CheckBox1_CheckedChanged
        RemoveHandler CheckBox2.CheckedChanged, AddressOf CheckBox2_CheckedChanged
        RemoveHandler CheckBox3.CheckedChanged, AddressOf CheckBox3_CheckedChanged
        RemoveHandler CheckBox4.CheckedChanged, AddressOf CheckBox4_CheckedChanged
        RemoveHandler CheckBox5.CheckedChanged, AddressOf CheckBox5_CheckedChanged
        RemoveHandler CheckBox6.CheckedChanged, AddressOf CheckBox6_CheckedChanged
        RemoveHandler CheckBox7.CheckedChanged, AddressOf CheckBox7_CheckedChanged
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_8()
        AddHandler Me.CheckBox7.CheckedChanged, AddressOf CheckBox7_CheckedChanged
        AddHandler Me.CheckBox6.CheckedChanged, AddressOf CheckBox6_CheckedChanged
        AddHandler Me.CheckBox5.CheckedChanged, AddressOf CheckBox5_CheckedChanged
        AddHandler Me.CheckBox4.CheckedChanged, AddressOf CheckBox4_CheckedChanged
        AddHandler Me.CheckBox3.CheckedChanged, AddressOf CheckBox3_CheckedChanged
        AddHandler Me.CheckBox2.CheckedChanged, AddressOf CheckBox2_CheckedChanged
        AddHandler Me.CheckBox1.CheckedChanged, AddressOf CheckBox1_CheckedChanged
        AddHandler Me.ComboBox8.SelectedIndexChanged, AddressOf ComboBox8_SelectedIndexChanged
        Call ARMAR_CADENA_NIVELES()
        If isuID_TIPO_USUARIO = 5 Then
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        Else
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        End If
        Call CARGAR_DATAGRIDVIEW_3()
        Call ARMAR_DATAGRIDVIEW_3()
    End Sub
    Private Sub ComboBox8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox8.SelectedIndexChanged
        Call ARMAR_CADENA_NIVELES()
        If isuID_TIPO_USUARIO = 5 Then
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        Else
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        End If
        Call CARGAR_DATAGRIDVIEW_3()
        Call ARMAR_DATAGRIDVIEW_3()
    End Sub
    Private Sub ARMAR_CADENA_NIVELES()
        CADENA_NIVELES = ""
        CADENA_NIVELES = "(ID_N1 = " & Me.ComboBox1.SelectedValue & ")"
        If Me.CheckBox1.Checked = True Then
            CADENA_NIVELES = CADENA_NIVELES & " AND (ID_N2 = " & Me.ComboBox2.SelectedValue & ")"
        Else
            GoTo SALIR
        End If
        If Me.CheckBox2.Checked = True Then
            CADENA_NIVELES = CADENA_NIVELES & " AND (ID_N3 = " & Me.ComboBox3.SelectedValue & ")"
        Else
            GoTo SALIR
        End If
        If Me.CheckBox3.Checked = True Then
            CADENA_NIVELES = CADENA_NIVELES & " AND (ID_N4 = " & Me.ComboBox4.SelectedValue & ")"
        Else
            GoTo SALIR
        End If
        If Me.CheckBox4.Checked = True Then
            CADENA_NIVELES = CADENA_NIVELES & " AND (ID_N5 = " & Me.ComboBox5.SelectedValue & ")"
        Else
            GoTo SALIR
        End If
        If Me.CheckBox5.Checked = True Then
            CADENA_NIVELES = CADENA_NIVELES & " AND (ID_N6 = " & Me.ComboBox6.SelectedValue & ")"
        Else
            GoTo SALIR
        End If
        If Me.CheckBox6.Checked = True Then
            CADENA_NIVELES = CADENA_NIVELES & " AND (ID_N7 = " & Me.ComboBox7.SelectedValue & ")"
        Else
            GoTo SALIR
        End If
        If Me.CheckBox7.Checked = True Then
            CADENA_NIVELES = CADENA_NIVELES & " AND (ID_N8 = " & Me.ComboBox8.SelectedValue & ")"
        Else
            GoTo SALIR
        End If
SALIR:
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_3()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql3, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE CARGOS SIT]")
        DGV3.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE CARGOS SIT]")
        Me.Label4.Text = DGV3.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_3()
        Me.DGV3.Columns(0).HeaderText = "Id"
        Me.DGV3.Columns(0).Width = 20
        Me.DGV3.Columns(0).Visible = True
        Me.DGV3.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV3.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Me.DGV3.Columns(1).HeaderText = "ORDEN"
        Me.DGV3.Columns(1).Width = 0
        Me.DGV3.Columns(1).Visible = False

        Me.DGV3.Columns(2).HeaderText = "ID_N1"
        Me.DGV3.Columns(2).Width = 0
        Me.DGV3.Columns(2).Visible = False

        Me.DGV3.Columns(3).HeaderText = "ORDEN_N1"
        Me.DGV3.Columns(3).Width = 0
        Me.DGV3.Columns(3).Visible = False

        Me.DGV3.Columns(4).HeaderText = "Nivel 1"
        Me.DGV3.Columns(4).Width = 0
        Me.DGV3.Columns(4).Visible = False
        Me.DGV3.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(5).HeaderText = "ID_N2"
        Me.DGV3.Columns(5).Width = 0
        Me.DGV3.Columns(5).Visible = False

        Me.DGV3.Columns(6).HeaderText = "ORDEN_N2"
        Me.DGV3.Columns(6).Width = 0
        Me.DGV3.Columns(6).Visible = False

        Me.DGV3.Columns(7).HeaderText = "Nivel 2"
        Me.DGV3.Columns(7).Width = 0
        Me.DGV3.Columns(7).Visible = False
        Me.DGV3.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(8).HeaderText = "ID_N3"
        Me.DGV3.Columns(8).Width = 0
        Me.DGV3.Columns(8).Visible = False

        Me.DGV3.Columns(9).HeaderText = "ORDEN_N3"
        Me.DGV3.Columns(9).Width = 0
        Me.DGV3.Columns(9).Visible = False

        Me.DGV3.Columns(10).HeaderText = "Nivel 3"
        Me.DGV3.Columns(10).Width = 0
        Me.DGV3.Columns(10).Visible = False
        Me.DGV3.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(11).HeaderText = "ID_N4"
        Me.DGV3.Columns(11).Width = 0
        Me.DGV3.Columns(11).Visible = False

        Me.DGV3.Columns(12).HeaderText = "ORDEN_N4"
        Me.DGV3.Columns(12).Width = 0
        Me.DGV3.Columns(12).Visible = False

        Me.DGV3.Columns(13).HeaderText = "Nivel 4"
        Me.DGV3.Columns(13).Width = 0
        Me.DGV3.Columns(13).Visible = False
        Me.DGV3.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(14).HeaderText = "ID_N5"
        Me.DGV3.Columns(14).Width = 0
        Me.DGV3.Columns(14).Visible = False

        Me.DGV3.Columns(15).HeaderText = "ORDEN_N5"
        Me.DGV3.Columns(15).Width = 0
        Me.DGV3.Columns(15).Visible = False

        Me.DGV3.Columns(16).HeaderText = "Nivel 5"
        Me.DGV3.Columns(16).Width = 0
        Me.DGV3.Columns(16).Visible = False
        Me.DGV3.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(17).HeaderText = "ID_N6"
        Me.DGV3.Columns(17).Width = 0
        Me.DGV3.Columns(17).Visible = False

        Me.DGV3.Columns(18).HeaderText = "ORDEN_N6"
        Me.DGV3.Columns(18).Width = 0
        Me.DGV3.Columns(18).Visible = False

        Me.DGV3.Columns(19).HeaderText = "Nivel 6"
        Me.DGV3.Columns(19).Width = 0
        Me.DGV3.Columns(19).Visible = False
        Me.DGV3.Columns(19).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(20).HeaderText = "ID_N7"
        Me.DGV3.Columns(20).Width = 0
        Me.DGV3.Columns(20).Visible = False

        Me.DGV3.Columns(21).HeaderText = "ORDEN_N7"
        Me.DGV3.Columns(21).Width = 0
        Me.DGV3.Columns(21).Visible = False

        Me.DGV3.Columns(22).HeaderText = "Nivel 7"
        Me.DGV3.Columns(22).Width = 0
        Me.DGV3.Columns(22).Visible = False
        Me.DGV3.Columns(22).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(23).HeaderText = "ID_N8"
        Me.DGV3.Columns(23).Width = 0
        Me.DGV3.Columns(23).Visible = False

        Me.DGV3.Columns(24).HeaderText = "ORDEN_N8"
        Me.DGV3.Columns(24).Width = 0
        Me.DGV3.Columns(24).Visible = False

        Me.DGV3.Columns(25).HeaderText = "Nivel 8"
        Me.DGV3.Columns(25).Width = 0
        Me.DGV3.Columns(25).Visible = False
        Me.DGV3.Columns(25).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(26).HeaderText = "ID_T_ES"
        Me.DGV3.Columns(26).Width = 0
        Me.DGV3.Columns(26).Visible = False

        Me.DGV3.Columns(27).HeaderText = "D_T_ES"
        Me.DGV3.Columns(27).Width = 0
        Me.DGV3.Columns(27).Visible = False

        Me.DGV3.Columns(28).HeaderText = "No. Orden [MIXTO]"
        Me.DGV3.Columns(28).Width = 50
        Me.DGV3.Columns(28).Visible = True
        Me.DGV3.Columns(28).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV3.Columns(28).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV3.Columns(28).DefaultCellStyle.BackColor = Color.Khaki

        Me.DGV3.Columns(29).HeaderText = "No. Orden [MILITAR]"
        Me.DGV3.Columns(29).Width = 50
        Me.DGV3.Columns(29).Visible = True
        Me.DGV3.Columns(29).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV3.Columns(29).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV3.Columns(29).DefaultCellStyle.BackColor = Color.LightGreen

        Me.DGV3.Columns(30).HeaderText = "No. Orden [PAME]"
        Me.DGV3.Columns(30).Width = 50
        Me.DGV3.Columns(30).Visible = True
        Me.DGV3.Columns(30).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV3.Columns(30).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV3.Columns(30).DefaultCellStyle.BackColor = Color.LightBlue

        Me.DGV3.Columns(31).HeaderText = "ID_CARGO_ES"
        Me.DGV3.Columns(31).Width = 0
        Me.DGV3.Columns(31).Visible = False

        Me.DGV3.Columns(32).HeaderText = "Cargo"
        Me.DGV3.Columns(32).Width = 180
        Me.DGV3.Columns(32).Visible = True
        Me.DGV3.Columns(32).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(33).HeaderText = "ID_GP"
        Me.DGV3.Columns(33).Width = 0
        Me.DGV3.Columns(33).Visible = False

        Me.DGV3.Columns(34).HeaderText = "Grado Plantilla"
        Me.DGV3.Columns(34).Width = 80
        Me.DGV3.Columns(34).Visible = True
        Me.DGV3.Columns(34).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(35).HeaderText = "ID_GR"
        Me.DGV3.Columns(35).Width = 0
        Me.DGV3.Columns(35).Visible = False

        Me.DGV3.Columns(36).HeaderText = "Grado Real"
        Me.DGV3.Columns(36).Width = 80
        Me.DGV3.Columns(36).Visible = True
        Me.DGV3.Columns(36).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(37).HeaderText = "ID_CAT_SALARIAL"
        Me.DGV3.Columns(37).Width = 0
        Me.DGV3.Columns(37).Visible = False

        Me.DGV3.Columns(38).HeaderText = "D_CAT_SALARIAL"
        Me.DGV3.Columns(38).Width = 0
        Me.DGV3.Columns(38).Visible = False

        Me.DGV3.Columns(39).HeaderText = "ID_M_P"
        Me.DGV3.Columns(39).Width = 0
        Me.DGV3.Columns(39).Visible = False

        Me.DGV3.Columns(40).HeaderText = "Codigo"
        Me.DGV3.Columns(40).Width = 80
        Me.DGV3.Columns(40).Visible = True
        Me.DGV3.Columns(40).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(41).HeaderText = "Apellidos y Nombres"
        Me.DGV3.Columns(41).Width = 220
        Me.DGV3.Columns(41).Visible = True
        Me.DGV3.Columns(41).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV3.Columns(41).DefaultCellStyle.BackColor = Color.Khaki

        Me.DGV3.Columns(42).HeaderText = "Nombres y Apellidos"
        Me.DGV3.Columns(42).Width = 0
        Me.DGV3.Columns(42).Visible = False

        Me.DGV3.Columns(43).HeaderText = "ID_SITUACION"
        Me.DGV3.Columns(43).Width = 0
        Me.DGV3.Columns(43).Visible = False

        Me.DGV3.Columns(44).HeaderText = "Situacion"
        Me.DGV3.Columns(44).Width = 60
        Me.DGV3.Columns(44).Visible = True
        Me.DGV3.Columns(44).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(45).HeaderText = "Jefe"
        Me.DGV3.Columns(45).Width = 40
        Me.DGV3.Columns(45).Visible = True
        Me.DGV3.Columns(45).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(46).HeaderText = "Mixta"
        Me.DGV3.Columns(46).Width = 0
        Me.DGV3.Columns(46).Visible = False
        Me.DGV3.Columns(46).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(47).HeaderText = "Militar"
        Me.DGV3.Columns(47).Width = 40
        Me.DGV3.Columns(47).Visible = True
        Me.DGV3.Columns(47).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(48).HeaderText = "Pame"
        Me.DGV3.Columns(48).Width = 40
        Me.DGV3.Columns(48).Visible = True
        Me.DGV3.Columns(48).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(49).HeaderText = "ID_ESTABLECIMIENTO"
        Me.DGV3.Columns(49).Width = 0
        Me.DGV3.Columns(49).Visible = False

        Me.DGV3.Columns(50).HeaderText = "Establecimiento"
        Me.DGV3.Columns(50).Width = 120
        Me.DGV3.Columns(50).Visible = True
        Me.DGV3.Columns(50).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(51).HeaderText = "D_ADMIN"
        Me.DGV3.Columns(51).Width = 0
        Me.DGV3.Columns(51).Visible = False
    End Sub
    Private Sub DGV3_RowPrePaint(sender As Object, e As DataGridViewRowPrePaintEventArgs) Handles DGV3.RowPrePaint
        Call FORMATO_FILAS_DATAGRIDVIEW()
    End Sub
    Private Sub FORMATO_FILAS_DATAGRIDVIEW()
        Dim style As New DataGridViewCellStyle
        style.Font = New Font(DGV3.Font, FontStyle.Bold)
        For Each Row As DataGridViewRow In DGV3.Rows
            If Row.Cells(50).Value = "VACANTE" Then
                Row.DefaultCellStyle.ForeColor = Color.Blue
            End If
            If Row.Cells(50).Value = "POLICLINICO ALTAMIRA" Then
                Row.DefaultCellStyle.BackColor = Color.LightGreen
            End If
            If Row.Cells(50).Value = "POLICLINICO TIPITAPA" Then
                Row.DefaultCellStyle.BackColor = Color.LightSkyBlue
            End If
            If Row.Cells(50).Value = "DIRECTORIO MEDICO PRIVADO" Then
                Row.DefaultCellStyle.BackColor = Color.PeachPuff
            End If
        Next
    End Sub
    Private Sub DGV3_Click(sender As Object, e As EventArgs) Handles DGV3.Click
        If Me.DGV3.RowCount > 0 Then
            For I = 0 To Me.DGV3.RowCount - 1
                If DGV3.Rows(I).Selected = True Then
                    Dim FILA = DGV3.CurrentRow.Index
                    If Not IsDBNull(Me.DGV3.Item(0, FILA).Value) Then : vmcsID_M_Ct = Me.DGV3.Item(0, FILA).Value : Else : vmcsID_M_Ct = 0 : End If     '*
                    If Not IsDBNull(Me.DGV3.Item(1, FILA).Value) Then : vmcsORDENt = Me.DGV3.Item(1, FILA).Value : Else : vmcsORDENt = "" : End If      '*
                    '----
                    If Not IsDBNull(Me.DGV3.Item(0, FILA).Value) Then : vmcsID_M_C = Me.DGV3.Item(0, FILA).Value : Else : vmcsID_M_C = 0 : End If
                    If Not IsDBNull(Me.DGV3.Item(1, FILA).Value) Then : vmcsORDEN = Me.DGV3.Item(1, FILA).Value : Else : vmcsORDEN = "" : End If
                    If Not IsDBNull(Me.DGV3.Item(2, FILA).Value) Then : vmcsID_N1 = Me.DGV3.Item(2, FILA).Value : Else : vmcsID_N1 = 0 : End If
                    If Not IsDBNull(Me.DGV3.Item(3, FILA).Value) Then : vmcsORDEN_N1 = Me.DGV3.Item(3, FILA).Value : Else : vmcsORDEN_N1 = "" : End If
                    If Not IsDBNull(Me.DGV3.Item(4, FILA).Value) Then : vmcsD_N1 = Me.DGV3.Item(4, FILA).Value : Else : vmcsD_N1 = "" : End If
                    If Not IsDBNull(Me.DGV3.Item(5, FILA).Value) Then : vmcsID_N2 = Me.DGV3.Item(5, FILA).Value : Else : vmcsID_N2 = 0 : End If
                    If Not IsDBNull(Me.DGV3.Item(6, FILA).Value) Then : vmcsORDEN_N2 = Me.DGV3.Item(6, FILA).Value : Else : vmcsORDEN_N2 = "" : End If
                    If Not IsDBNull(Me.DGV3.Item(7, FILA).Value) Then : vmcsD_N2 = Me.DGV3.Item(7, FILA).Value : Else : vmcsD_N2 = "" : End If
                    If Not IsDBNull(Me.DGV3.Item(8, FILA).Value) Then : vmcsID_N3 = Me.DGV3.Item(8, FILA).Value : Else : vmcsID_N3 = 0 : End If
                    If Not IsDBNull(Me.DGV3.Item(9, FILA).Value) Then : vmcsORDEN_N3 = Me.DGV3.Item(9, FILA).Value : Else : vmcsORDEN_N3 = "" : End If
                    If Not IsDBNull(Me.DGV3.Item(10, FILA).Value) Then : vmcsD_N3 = Me.DGV3.Item(10, FILA).Value : Else : vmcsD_N3 = "" : End If
                    If Not IsDBNull(Me.DGV3.Item(11, FILA).Value) Then : vmcsID_N4 = Me.DGV3.Item(11, FILA).Value : Else : vmcsID_N4 = 0 : End If
                    If Not IsDBNull(Me.DGV3.Item(12, FILA).Value) Then : vmcsORDEN_N4 = Me.DGV3.Item(12, FILA).Value : Else : vmcsORDEN_N4 = "" : End If
                    If Not IsDBNull(Me.DGV3.Item(13, FILA).Value) Then : vmcsD_N4 = Me.DGV3.Item(13, FILA).Value : Else : vmcsD_N4 = "" : End If
                    If Not IsDBNull(Me.DGV3.Item(14, FILA).Value) Then : vmcsID_N5 = Me.DGV3.Item(14, FILA).Value : Else : vmcsID_N5 = 0 : End If
                    If Not IsDBNull(Me.DGV3.Item(15, FILA).Value) Then : vmcsORDEN_N5 = Me.DGV3.Item(15, FILA).Value : Else : vmcsORDEN_N5 = "" : End If
                    If Not IsDBNull(Me.DGV3.Item(16, FILA).Value) Then : vmcsD_N5 = Me.DGV3.Item(16, FILA).Value : Else : vmcsD_N5 = "" : End If
                    If Not IsDBNull(Me.DGV3.Item(17, FILA).Value) Then : vmcsID_N6 = Me.DGV3.Item(17, FILA).Value : Else : vmcsID_N6 = 0 : End If
                    If Not IsDBNull(Me.DGV3.Item(18, FILA).Value) Then : vmcsORDEN_N6 = Me.DGV3.Item(18, FILA).Value : Else : vmcsORDEN_N6 = "" : End If
                    If Not IsDBNull(Me.DGV3.Item(19, FILA).Value) Then : vmcsD_N6 = Me.DGV3.Item(19, FILA).Value : Else : vmcsD_N6 = "" : End If
                    If Not IsDBNull(Me.DGV3.Item(20, FILA).Value) Then : vmcsID_N7 = Me.DGV3.Item(20, FILA).Value : Else : vmcsID_N7 = 0 : End If
                    If Not IsDBNull(Me.DGV3.Item(21, FILA).Value) Then : vmcsORDEN_N7 = Me.DGV3.Item(21, FILA).Value : Else : vmcsORDEN_N7 = "" : End If
                    If Not IsDBNull(Me.DGV3.Item(22, FILA).Value) Then : vmcsD_N7 = Me.DGV3.Item(22, FILA).Value : Else : vmcsD_N7 = "" : End If
                    If Not IsDBNull(Me.DGV3.Item(23, FILA).Value) Then : vmcsID_N8 = Me.DGV3.Item(23, FILA).Value : Else : vmcsID_N8 = 0 : End If
                    If Not IsDBNull(Me.DGV3.Item(24, FILA).Value) Then : vmcsORDEN_N8 = Me.DGV3.Item(24, FILA).Value : Else : vmcsORDEN_N8 = "" : End If
                    If Not IsDBNull(Me.DGV3.Item(25, FILA).Value) Then : vmcsD_N8 = Me.DGV3.Item(25, FILA).Value : Else : vmcsD_N8 = "" : End If
                    If Not IsDBNull(Me.DGV3.Item(26, FILA).Value) Then : vmcsID_T_ES = Me.DGV3.Item(26, FILA).Value : Else : vmcsID_T_ES = 0 : End If
                    If Not IsDBNull(Me.DGV3.Item(27, FILA).Value) Then : vmcsD_T_ES = Me.DGV3.Item(27, FILA).Value : Else : vmcsD_T_ES = "" : End If
                    If Not IsDBNull(Me.DGV3.Item(28, FILA).Value) Then : vmcsN_ORDEN_MIXTA = Me.DGV3.Item(28, FILA).Value : Else : vmcsN_ORDEN_MIXTA = "" : End If
                    If Not IsDBNull(Me.DGV3.Item(29, FILA).Value) Then : vmcsN_ORDEN_MILITAR = Me.DGV3.Item(29, FILA).Value : Else : vmcsN_ORDEN_MILITAR = "" : End If
                    If Not IsDBNull(Me.DGV3.Item(30, FILA).Value) Then : vmcsN_ORDEN_PAME = Me.DGV3.Item(30, FILA).Value : Else : vmcsN_ORDEN_PAME = "" : End If
                    If Not IsDBNull(Me.DGV3.Item(31, FILA).Value) Then : vmcsID_CARGO_ES = Me.DGV3.Item(31, FILA).Value : Else : vmcsID_CARGO_ES = 0 : End If
                    If Not IsDBNull(Me.DGV3.Item(32, FILA).Value) Then : vmcsD_CARGO_ES = Me.DGV3.Item(32, FILA).Value : Else : vmcsD_CARGO_ES = "" : End If
                    If Not IsDBNull(Me.DGV3.Item(33, FILA).Value) Then : vmcsID_GP = Me.DGV3.Item(33, FILA).Value : Else : vmcsID_GP = 0 : End If
                    If Not IsDBNull(Me.DGV3.Item(34, FILA).Value) Then : vmcsD_GP = Me.DGV3.Item(34, FILA).Value : Else : vmcsD_GP = "" : End If
                    If Not IsDBNull(Me.DGV3.Item(35, FILA).Value) Then : vmcsID_GR = Me.DGV3.Item(35, FILA).Value : Else : vmcsID_GR = 0 : End If
                    If Not IsDBNull(Me.DGV3.Item(36, FILA).Value) Then : vmcsD_GR = Me.DGV3.Item(36, FILA).Value : Else : vmcsD_GR = "" : End If
                    If Not IsDBNull(Me.DGV3.Item(37, FILA).Value) Then : vmcsID_CAT_SALARIAL = Me.DGV3.Item(37, FILA).Value : Else : vmcsID_CAT_SALARIAL = 0 : End If
                    If Not IsDBNull(Me.DGV3.Item(38, FILA).Value) Then : vmcsD_CAT_SALARIAL = Me.DGV3.Item(38, FILA).Value : Else : vmcsD_CAT_SALARIAL = "" : End If
                    If Not IsDBNull(Me.DGV3.Item(39, FILA).Value) Then : vmcsID_M_P = Me.DGV3.Item(39, FILA).Value : Else : vmcsID_M_P = 0 : End If
                    If Not IsDBNull(Me.DGV3.Item(40, FILA).Value) Then : vmcsCODIGO = Me.DGV3.Item(40, FILA).Value : Else : vmcsCODIGO = "" : End If
                    If Not IsDBNull(Me.DGV3.Item(41, FILA).Value) Then : vmcsAN = Me.DGV3.Item(41, FILA).Value : Else : vmcsAN = "" : End If
                    If Not IsDBNull(Me.DGV3.Item(42, FILA).Value) Then : vmcsNA = Me.DGV3.Item(42, FILA).Value : Else : vmcsNA = "" : End If
                    If Not IsDBNull(Me.DGV3.Item(43, FILA).Value) Then : vmcsID_SITUACION = Me.DGV3.Item(43, FILA).Value : Else : vmcsID_SITUACION = 0 : End If
                    If Not IsDBNull(Me.DGV3.Item(44, FILA).Value) Then : vmcsD_SITUACION = Me.DGV3.Item(44, FILA).Value : Else : vmcsD_SITUACION = "" : End If
                    If Not IsDBNull(Me.DGV3.Item(45, FILA).Value) Then : vmcsJEFE = Me.DGV3.Item(45, FILA).Value : Else : vmcsJEFE = False : End If
                    If Not IsDBNull(Me.DGV3.Item(46, FILA).Value) Then : vmcsMIXTA = Me.DGV3.Item(46, FILA).Value : Else : vmcsMIXTA = False : End If
                    If Not IsDBNull(Me.DGV3.Item(47, FILA).Value) Then : vmcsMILITAR = Me.DGV3.Item(47, FILA).Value : Else : vmcsMILITAR = False : End If
                    If Not IsDBNull(Me.DGV3.Item(48, FILA).Value) Then : vmcsPAME = Me.DGV3.Item(48, FILA).Value : Else : vmcsPAME = False : End If
                    If Not IsDBNull(Me.DGV3.Item(49, FILA).Value) Then : vmcsID_ESTABLECIMIENTO = Me.DGV3.Item(49, FILA).Value : Else : vmcsID_ESTABLECIMIENTO = 0 : End If
                    If Not IsDBNull(Me.DGV3.Item(50, FILA).Value) Then : vmcsD_ESTABLECIMIENTO = Me.DGV3.Item(50, FILA).Value : Else : vmcsD_ESTABLECIMIENTO = "" : End If
                    Exit Sub
                End If
            Next
        End If
    End Sub
    'Private Sub DGV3_SelectionChanged(sender As Object, e As EventArgs) Handles DGV3.SelectionChanged
    '    If Me.DGV3.RowCount > 0 Then
    '        For I = 0 To Me.DGV3.RowCount - 1
    '            If DGV3.Rows(I).Selected = True Then
    '                vmcsID_M_Ct = Me.DGV3.Item(0, I).Value  '*
    '                vmcsORDENt = Me.DGV3.Item(1, I).Value   '*
    '                '----
    '                vmcsID_M_C = Me.DGV3.Item(0, I).Value
    '                vmcsORDEN = Me.DGV3.Item(1, I).Value
    '                vmcsID_N1 = Me.DGV3.Item(2, I).Value
    '                vmcsORDEN_N1 = Me.DGV3.Item(3, I).Value
    '                vmcsD_N1 = Me.DGV3.Item(4, I).Value
    '                vmcsID_N2 = Me.DGV3.Item(5, I).Value
    '                vmcsORDEN_N2 = Me.DGV3.Item(6, I).Value
    '                vmcsD_N2 = Me.DGV3.Item(7, I).Value
    '                vmcsID_T_ES = Me.DGV3.Item(8, I).Value
    '                vmcsD_T_ES = Me.DGV3.Item(9, I).Value
    '                vmcsN_ORDEN_MIXTA = Me.DGV3.Item(10, I).Value
    '                vmcsN_ORDEN_MILITAR = Me.DGV3.Item(11, I).Value
    '                vmcsN_ORDEN_PAME = Me.DGV3.Item(12, I).Value
    '                vmcsID_CARGO_ES = Me.DGV3.Item(13, I).Value
    '                vmcsD_CARGO_ES = Me.DGV3.Item(14, I).Value
    '                vmcsID_GP = Me.DGV3.Item(15, I).Value
    '                vmcsD_GP = Me.DGV3.Item(16, I).Value
    '                vmcsID_GR = Me.DGV3.Item(17, I).Value
    '                vmcsD_GR = Me.DGV3.Item(18, I).Value
    '                vmcsID_CAT_SALARIAL = Me.DGV3.Item(19, I).Value
    '                vmcsD_CAT_SALARIAL = Me.DGV3.Item(20, I).Value
    '                vmcsID_M_P = Me.DGV3.Item(21, I).Value
    '                vmcsCODIGO = Me.DGV3.Item(22, I).Value
    '                vmcsAN = Me.DGV3.Item(23, I).Value
    '                vmcsID_SITUACION = Me.DGV3.Item(24, I).Value
    '                vmcsD_SITUACION = Me.DGV3.Item(25, I).Value
    '                vmcsJEFE = Me.DGV3.Item(26, I).Value
    '                vmcsMIXTA = Me.DGV3.Item(27, I).Value
    '                vmcsMILITAR = Me.DGV3.Item(28, I).Value
    '                vmcsPAME = Me.DGV3.Item(29, I).Value
    '                Exit Sub
    '            End If
    '        Next
    '    End If
    'End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click   'SUBIR
        Call VERIFICAR_ACCESOS("057") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmcsID_M_Ct = 0 Then
            MsgBox("Debe seleccionar o hacer clic sobre el Registro/ Cargo que desea mover", vbInformation, "Mensaje del Sistema")
            Me.DGV3.Focus()
            Exit Sub
        End If
        vmcsORDENt = vmcsORDENt - 1.5
        Call ACTUALIZAR_ORDEN_DE_CARGO_EN_MAESTRO_DE_CARGOS()
        Call REENUMERAR_ORDEN()
        If isuID_TIPO_USUARIO = 5 Then
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        Else
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        End If
        Call CARGAR_DATAGRIDVIEW_3()
        Call ARMAR_DATAGRIDVIEW_3()
        Call SELECCIONAR_DATAGRIDVIEW_3()
    End Sub
    Private Sub ACTUALIZAR_ORDEN_DE_CARGO_EN_MAESTRO_DE_CARGOS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE CARGOS] SET ORDEN = '" & vmcsORDENt & "' WHERE ID_M_C = " & vmcsID_M_Ct & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub SELECCIONAR_DATAGRIDVIEW_3()
        Dim I As Integer
        For I = 0 To Me.DGV3.RowCount - 1
            If Me.DGV3.Item(0, I).Value = vmcsID_M_Ct Then
                Me.DGV3.Rows(I).Selected = True
                Me.DGV3.CurrentCell = Me.DGV3.Item(0, I)
                vmcsORDENt = Me.DGV3.Item(1, I).Value
                Exit Sub
            End If
        Next
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click   'BAJAR
        Call VERIFICAR_ACCESOS("058") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmcsID_M_Ct = 0 Then
            MsgBox("Debe seleccionar o hacer clic sobre el Registro/ Cargo que desea mover", vbInformation, "Mensaje del Sistema")
            Me.DGV3.Focus()
            Exit Sub
        End If
        vmcsORDENt = vmcsORDENt + 1.5
        Call ACTUALIZAR_ORDEN_DE_CARGO_EN_MAESTRO_DE_CARGOS()
        Call REENUMERAR_ORDEN()
        If isuID_TIPO_USUARIO = 5 Then
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        Else
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        End If
        Call CARGAR_DATAGRIDVIEW_3()
        Call ARMAR_DATAGRIDVIEW_3()
        Call SELECCIONAR_DATAGRIDVIEW_3()
    End Sub
    Dim val_ID_M_C As Integer
    Dim New_N_ORDEN As Integer
    Private Sub REENUMERAR_ORDEN()
        New_N_ORDEN = 0
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [VISTA MAESTRO DE CARGOS SIT] ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                For I = 0 To MiDataTable.Rows.Count - 1
                    val_ID_M_C = MiDataTable.Rows(I).Item(0).ToString
                    New_N_ORDEN += 1
                    Call ACTUALIZAR_New_N_ORDEN()
                Next
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ACTUALIZAR_New_N_ORDEN()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE CARGOS] SET ORDEN = '" & New_N_ORDEN & "' WHERE ID_M_C = " & val_ID_M_C & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click               'RENUMERAR
        Call VERIFICAR_ACCESOS("059") : If HAY_ACCESO = False Then : Exit Sub : End If
        MENSAJE = MsgBox("Se Dispone a Actualizar los Números de Orden, ¿Esta seguro de continuar?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            Call REENUMERAR()
            MsgBox("Proceso Finalizado Satisfactoriamente", vbInformation, "Mensaje del Sistema")
        End If
    End Sub
    Dim bus_ID_M_C, New_ORDEN As Integer
    Private Sub REENUMERAR()
        bus_ID_M_C = 0
        New_ORDEN = 0
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [dbo].[VISTA MAESTRO DE CARGOS SIT] ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                For I = 0 To MiDataTable.Rows.Count - 1
                    bus_ID_M_C = MiDataTable.Rows(I).Item(0).ToString
                    New_ORDEN += 1
                    Call ACTUALIZAR_ORDEN()
                Next
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click               'AGREGAR
        Call VERIFICAR_ACCESOS("051") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmcsID_M_C = 0 Then
            MsgBox("Debe seleccionar la Ubicacion del Cargo en el Servicio o Departamento", vbInformation, "Mensaje del Sistema")
            vmcsID_M_C = 0
            Me.DGV3.Focus()
            Exit Sub
        End If
        Frm_Estructura_Org_Agregar.ShowDialog()
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click               'ELIMINAR
        Call VERIFICAR_ACCESOS("053") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmcsID_M_C = 0 Then
            MsgBox("Debe seleccionar el registro que desea Eliminar", vbInformation, "Mensaje del Sistema")
            Me.DGV3.Focus()
            Exit Sub
        End If
        If vmcsID_SITUACION = 1 Then
            MsgBox("No es posible Eliminar el cargo seleccionado, cargo OCUPADO\ CONGELADO", vbInformation, "Mensaje del Sistema")
            Me.DGV3.Focus()
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Esta seguro de Eliminar el registro seleccionado?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            Call ELIMINAR_MAESTRO_DE_CARGOS()
            If isuID_TIPO_USUARIO = 5 Then
                CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
            Else
                CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
            End If
            Call CARGAR_DATAGRIDVIEW_3()
            Call ARMAR_DATAGRIDVIEW_3()
        End If
    End Sub
    Private Sub ELIMINAR_MAESTRO_DE_CARGOS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE CARGOS] WHERE ID_M_C = " & CInt(vmcsID_M_C) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click                   'EDITAR
        Call VERIFICAR_ACCESOS("052") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmcsID_M_C = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Me.DGV3.Focus()
            Exit Sub
        End If
        Frm_Estructura_Org_Editar.ShowDialog()
    End Sub
    Private Sub ACTUALIZAR_ORDEN()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE CARGOS] SET ORDEN = '" & New_ORDEN & "', N_ORDEN_MIXTA = '" & Format(New_ORDEN, "0000") & "', N_ORDEN_PAME = '" & Format(New_ORDEN, "0000") & "' WHERE ID_M_C = " & bus_ID_M_C & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click                   'TRASLADAR
        Call VERIFICAR_ACCESOS("055") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Estructura_Org_Trasladar.ShowDialog()
        If isuID_TIPO_USUARIO = 5 Then
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        Else
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        End If
        Call CARGAR_DATAGRIDVIEW_3()
        Call ARMAR_DATAGRIDVIEW_3()
    End Sub
    Private Sub DGV3_DoubleClick(sender As Object, e As EventArgs) Handles DGV3.DoubleClick             'EDITAR
        Call VERIFICAR_ACCESOS("052") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmcsID_M_C = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Me.DGV3.Focus()
            Exit Sub
        End If
        Frm_Estructura_Org_Editar.ShowDialog()
    End Sub
    Private Sub Frm_Estructura_Org_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        ESTRUCTURA = False
    End Sub
    Private Sub Frm_Estructura_Org_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        ESTRUCTURA = False
    End Sub
    Private Sub Frm_Estructura_Org_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        ESTRUCTURA = False
    End Sub
    Private Sub Frm_Estructura_Org_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        ESTRUCTURA = False
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If Me.CheckBox1.Checked = True Then
            Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_2()
            NIVEL2sql = Me.ComboBox2.SelectedValue
            Call ARMAR_CADENA_NIVELES()
        Else
            Me.ComboBox2.DataSource = Nothing
            Me.ComboBox2.ForeColor = Color.Red
            Me.ComboBox2.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
            Me.ComboBox2.Text = "-- SIN DATOS -- "
            NIVEL2sql = Nothing
            Me.CheckBox2.Checked = False
            Me.CheckBox3.Checked = False
            Me.CheckBox4.Checked = False
            Call ARMAR_CADENA_NIVELES()
        End If
        If isuID_TIPO_USUARIO = 5 Then
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        Else
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        End If
        Call CARGAR_DATAGRIDVIEW_3()
        Call ARMAR_DATAGRIDVIEW_3()
    End Sub
    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If Me.CheckBox2.Checked = True Then
            Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_3()
            NIVEL3sql = Me.ComboBox3.SelectedValue
            Call ARMAR_CADENA_NIVELES()
        Else
            Me.ComboBox3.DataSource = Nothing
            Me.ComboBox3.ForeColor = Color.Red
            Me.ComboBox3.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
            Me.ComboBox3.Text = "-- SIN DATOS -- "
            NIVEL3sql = Nothing
            Me.CheckBox3.Checked = False
            Me.CheckBox4.Checked = False
            Call ARMAR_CADENA_NIVELES()
        End If
        If isuID_TIPO_USUARIO = 5 Then
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        Else
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        End If
        Call CARGAR_DATAGRIDVIEW_3()
        Call ARMAR_DATAGRIDVIEW_3()
    End Sub
    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If Me.CheckBox3.Checked = True Then
            Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_4()
            NIVEL4sql = Me.ComboBox4.SelectedValue
            Call ARMAR_CADENA_NIVELES()
        Else
            Me.ComboBox4.DataSource = Nothing
            Me.ComboBox4.ForeColor = Color.Red
            Me.ComboBox4.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
            Me.ComboBox4.Text = "-- SIN DATOS -- "
            NIVEL4sql = Nothing
            Me.CheckBox4.Checked = False
            Call ARMAR_CADENA_NIVELES()
        End If
        If isuID_TIPO_USUARIO = 5 Then
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        Else
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        End If
        Call CARGAR_DATAGRIDVIEW_3()
        Call ARMAR_DATAGRIDVIEW_3()
    End Sub
    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If Me.CheckBox4.Checked = True Then
            Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_5
            NIVEL5sql = Me.ComboBox5.SelectedValue
            Call ARMAR_CADENA_NIVELES()
        Else
            Me.ComboBox5.DataSource = Nothing
            Me.ComboBox5.ForeColor = Color.Red
            Me.ComboBox5.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
            Me.ComboBox5.Text = "-- SIN DATOS -- "
            NIVEL5sql = Nothing
            Call ARMAR_CADENA_NIVELES()
        End If
        If isuID_TIPO_USUARIO = 5 Then
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        Else
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        End If
        Call CARGAR_DATAGRIDVIEW_3()
        Call ARMAR_DATAGRIDVIEW_3()
    End Sub
    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If Me.CheckBox5.Checked = True Then
            Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_6()
            NIVEL6sql = Me.ComboBox6.SelectedValue
            Call ARMAR_CADENA_NIVELES()
        Else
            Me.ComboBox6.DataSource = Nothing
            Me.ComboBox6.ForeColor = Color.Red
            Me.ComboBox6.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
            Me.ComboBox6.Text = "-- SIN DATOS -- "
            NIVEL6sql = Nothing
            Call ARMAR_CADENA_NIVELES()
        End If
        If isuID_TIPO_USUARIO = 5 Then
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        Else
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        End If
        Call CARGAR_DATAGRIDVIEW_3()
        Call ARMAR_DATAGRIDVIEW_3()
    End Sub
    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        If Me.CheckBox6.Checked = True Then
            Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_7()
            NIVEL7sql = Me.ComboBox7.SelectedValue
            Call ARMAR_CADENA_NIVELES()
        Else
            Me.ComboBox7.DataSource = Nothing
            Me.ComboBox7.ForeColor = Color.Red
            Me.ComboBox7.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
            Me.ComboBox7.Text = "-- SIN DATOS -- "
            NIVEL7sql = Nothing
            Call ARMAR_CADENA_NIVELES()
        End If
        If isuID_TIPO_USUARIO = 5 Then
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        Else
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        End If
        Call CARGAR_DATAGRIDVIEW_3()
        Call ARMAR_DATAGRIDVIEW_3()
    End Sub
    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        If Me.CheckBox7.Checked = True Then
            Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_8()
            NIVEL8sql = Me.ComboBox8.SelectedValue
            Call ARMAR_CADENA_NIVELES()
        Else
            Me.ComboBox8.DataSource = Nothing
            Me.ComboBox8.ForeColor = Color.Red
            Me.ComboBox8.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
            Me.ComboBox8.Text = "-- SIN DATOS -- "
            NIVEL8sql = Nothing
            Call ARMAR_CADENA_NIVELES()
        End If
        If isuID_TIPO_USUARIO = 5 Then
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        Else
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        End If
        Call CARGAR_DATAGRIDVIEW_3()
        Call ARMAR_DATAGRIDVIEW_3()
    End Sub
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Call VERIFICAR_ACCESOS("054") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Estructura_Org_Imprimir.ShowDialog()
    End Sub
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Call VERIFICAR_ACCESOS("050") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Estructura_Org_Buscar.ShowDialog()
        If CERRAR = False Then
            'If cN1 = "-- SIN DATOS -- " Then
            '    Me.CheckBox1.Checked = False
            'Else
            '    Me.ComboBox1.Text = cN1
            'End If
            If cN2 = "-- SIN DATOS -- " Then
                Me.CheckBox1.Checked = False
            Else
                Me.CheckBox1.Checked = True
                Me.ComboBox2.Text = cN2
            End If
            If cN3 = "-- SIN DATOS -- " Then
                Me.CheckBox2.Checked = False
            Else
                Me.CheckBox2.Checked = True
                Me.ComboBox3.Text = cN3
            End If
            If cN4 = "-- SIN DATOS -- " Then
                Me.CheckBox3.Checked = False
            Else
                Me.CheckBox3.Checked = True
                Me.ComboBox4.Text = cN4
            End If
            If cN5 = "-- SIN DATOS -- " Then
                Me.CheckBox4.Checked = False
            Else
                Me.CheckBox4.Checked = True
                Me.ComboBox5.Text = cN5
            End If
            If cN6 = "-- SIN DATOS -- " Then
                Me.CheckBox5.Checked = False
            Else
                Me.CheckBox5.Checked = True
                Me.ComboBox6.Text = cN6
            End If
            If cN7 = "-- SIN DATOS -- " Then
                Me.CheckBox6.Checked = False
            Else
                Me.CheckBox6.Checked = True
                Me.ComboBox7.Text = cN7
            End If
            If cN8 = "-- SIN DATOS -- " Then
                Me.CheckBox7.Checked = False
            Else
                Me.CheckBox7.Checked = True
                Me.ComboBox8.Text = cN8
            End If

            'Me.ComboBox2.Text = cN2
            'Me.ComboBox3.Text = cN3
            'Me.ComboBox4.Text = cN4
            'Me.ComboBox5.Text = cN5
            'Me.ComboBox6.Text = cN6
            'Me.ComboBox7.Text = cN7
            'Me.ComboBox8.Text = cN8
            Call ARMAR_CADENA_NIVELES()
            If isuID_TIPO_USUARIO = 5 Then
                CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
            Else
                CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
            End If
            Call CARGAR_DATAGRIDVIEW_3()
            Call ARMAR_DATAGRIDVIEW_3()
            Call SELECCIONAR_DATAGRIDVIEW_3()
        End If
    End Sub
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Call VERIFICAR_ACCESOS("056") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Estructura_Org_Importar.ShowDialog()
        If CERRAR = False Then
            Call ARMAR_CADENA_NIVELES()
            If isuID_TIPO_USUARIO = 5 Then
                CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
            Else
                CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
            End If
            Call CARGAR_DATAGRIDVIEW_3()
            Call ARMAR_DATAGRIDVIEW_3()
        End If
    End Sub
End Class