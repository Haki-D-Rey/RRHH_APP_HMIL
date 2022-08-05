Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Public Class Frm_Consulta_Saldos
    Private Sub Frm_Consulta_Saldos_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            CONSULTA_SAL_VAC = False
            Me.Close()
        End If
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
        'If isuID_TIPO_USUARIO = 5 Then
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'Else
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'End If
        'Call CARGAR_DATAGRIDVIEW_3()
        'Call ARMAR_DATAGRIDVIEW_3()
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
        'If isuID_TIPO_USUARIO = 5 Then
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'Else
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'End If
        'Call CARGAR_DATAGRIDVIEW_3()
        'Call ARMAR_DATAGRIDVIEW_3()
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
        'If isuID_TIPO_USUARIO = 5 Then
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'Else
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'End If
        'Call CARGAR_DATAGRIDVIEW_3()
        'Call ARMAR_DATAGRIDVIEW_3()
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
        'If isuID_TIPO_USUARIO = 5 Then
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'Else
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'End If
        'Call CARGAR_DATAGRIDVIEW_3()
        'Call ARMAR_DATAGRIDVIEW_3()
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
        'If isuID_TIPO_USUARIO = 5 Then
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'Else
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'End If
        'Call CARGAR_DATAGRIDVIEW_3()
        'Call ARMAR_DATAGRIDVIEW_3()
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
        'If isuID_TIPO_USUARIO = 5 Then
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'Else
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'End If
        'Call CARGAR_DATAGRIDVIEW_3()
        'Call ARMAR_DATAGRIDVIEW_3()
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
        'If isuID_TIPO_USUARIO = 5 Then
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'Else
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'End If
        'Call CARGAR_DATAGRIDVIEW_3()
        'Call ARMAR_DATAGRIDVIEW_3()
    End Sub
    Private Sub ComboBox8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox8.SelectedIndexChanged
        Call ARMAR_CADENA_NIVELES()
        'If isuID_TIPO_USUARIO = 5 Then
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'Else
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'End If
        'Call CARGAR_DATAGRIDVIEW_3()
        'Call ARMAR_DATAGRIDVIEW_3()
    End Sub
    Private Sub ARMAR_CADENA_NIVELES()
        CADENA_NIVELES = ""
        CADENA_NIVELES = "([NIVEL 1] = '" & Me.ComboBox1.Text & "')"
        If Me.CheckBox1.Checked = True Then
            CADENA_NIVELES = CADENA_NIVELES & " AND ([NIVEL 2] = '" & Me.ComboBox2.Text & "')"
        Else
            GoTo SALIR
        End If
        If Me.CheckBox2.Checked = True Then
            CADENA_NIVELES = CADENA_NIVELES & " AND ([NIVEL 3] = '" & Me.ComboBox3.Text & "')"
        Else
            GoTo SALIR
        End If
        If Me.CheckBox3.Checked = True Then
            CADENA_NIVELES = CADENA_NIVELES & " AND ([NIVEL 4] = '" & Me.ComboBox4.Text & "')"
        Else
            GoTo SALIR
        End If
        If Me.CheckBox4.Checked = True Then
            CADENA_NIVELES = CADENA_NIVELES & " AND ([NIVEL 5] = '" & Me.ComboBox5.Text & "')"
        Else
            GoTo SALIR
        End If
        If Me.CheckBox5.Checked = True Then
            CADENA_NIVELES = CADENA_NIVELES & " AND ([NIVEL 6] = '" & Me.ComboBox6.Text & "')"
        Else
            GoTo SALIR
        End If
        If Me.CheckBox6.Checked = True Then
            CADENA_NIVELES = CADENA_NIVELES & " AND ([NIVEL 7] = '" & Me.ComboBox7.Text & "')"
        Else
            GoTo SALIR
        End If
        If Me.CheckBox7.Checked = True Then
            CADENA_NIVELES = CADENA_NIVELES & " AND ([NIVEL 8] = '" & Me.ComboBox8.Text & "')"
        Else
            GoTo SALIR
        End If
SALIR:
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
        'If isuID_TIPO_USUARIO = 5 Then
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'Else
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'End If
        'Call CARGAR_DATAGRIDVIEW_3()
        'Call ARMAR_DATAGRIDVIEW_3()
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
        'If isuID_TIPO_USUARIO = 5 Then
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'Else
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'End If
        'Call CARGAR_DATAGRIDVIEW_3()
        'Call ARMAR_DATAGRIDVIEW_3()
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
        'If isuID_TIPO_USUARIO = 5 Then
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'Else
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'End If
        'Call CARGAR_DATAGRIDVIEW_3()
        'Call ARMAR_DATAGRIDVIEW_3()
    End Sub
    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If Me.CheckBox4.Checked = True Then
            Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_5()
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
        'If isuID_TIPO_USUARIO = 5 Then
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'Else
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'End If
        'Call CARGAR_DATAGRIDVIEW_3()
        'Call ARMAR_DATAGRIDVIEW_3()
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
        'If isuID_TIPO_USUARIO = 5 Then
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'Else
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'End If
        'Call CARGAR_DATAGRIDVIEW_3()
        'Call ARMAR_DATAGRIDVIEW_3()
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
        'If isuID_TIPO_USUARIO = 5 Then
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'Else
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'End If
        'Call CARGAR_DATAGRIDVIEW_3()
        'Call ARMAR_DATAGRIDVIEW_3()
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
        'If isuID_TIPO_USUARIO = 5 Then
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'Else
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'End If
        'Call CARGAR_DATAGRIDVIEW_3()
        'Call ARMAR_DATAGRIDVIEW_3()
    End Sub
    Private Sub Frm_Consulta_Saldos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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

        Me.RadioButton2.Checked = True
        Me.DGV1.DataSource = Nothing
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        CONSULTA_SAL_VAC = False
        Me.Close()
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW()
        Application.DoEvents()
        Me.Cursor = Cursors.WaitCursor
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql1, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA CONSULTA GENERAL DE EXPEDIENTES SALDOS VAC]")
        DGV1.DataSource = MiDataSet.Tables("[VISTA CONSULTA GENERAL DE EXPEDIENTES SALDOS VAC]")
        Me.Label4.Text = DGV1.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Me.Cursor = Cursors.Default
    End Sub
    Dim CAMPOS As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call VERIFICAR_ACCESOS("303") : If HAY_ACCESO = False Then : Exit Sub : End If
        If Me.RadioButton1.Checked = False And Me.RadioButton2.Checked = False Then
            MsgBox("Debe seleccionar el Tipo de Estructura primeramente", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        SELECCION = ""
        If Me.RadioButton1.Checked = True Then
            SELECCION = "[CARGO MILITAR] = 1"
        End If
        If Me.RadioButton2.Checked = True Then
            SELECCION = "[CARGO PAME] = 1"
        End If
        'If Me.CheckBox8.Checked = True Then
        CADENAsql1 = "SELECT ORDEN, [NIVEL 1], [NIVEL 2], [NIVEL 3], [NIVEL 4], [NIVEL 5], [NIVEL 6], [FECHA INGRESO PAME], [CODIGO EMPLEADO], APELLIDOS, NOMBRES, [CARGO MILITAR], [CARGO PAME], [SALDO VACACIONAL ACUMULADO], ACREDITA, [VACACIONES CANCELADAS], [OMISIONES DE MARCAS], [AUSENCIA INJUSTIFICADA], AMP, AMN, SALDO FROM [VISTA CONSULTA GENERAL DE EXPEDIENTES SALDOS VAC] WHERE " & CADENA_NIVELES & " and " & SELECCION & ""
        'Else
        'CADENAsql1 = "SELECT ORDEN, dbo.ProperCase([NIVEL 1]) AS [NIVEL 1], dbo.ProperCase([NIVEL 2]) AS [NIVEL 2], dbo.ProperCase([NIVEL 3]) AS [NIVEL 3], dbo.ProperCase([NIVEL 4]) AS [NIVEL 4], dbo.ProperCase([NIVEL 5]) AS [NIVEL 5], dbo.ProperCase([NIVEL 6]) AS [NIVEL 6], dbo.ProperCase([FECHA INGRESO PAME]) AS [FECHA INGRESO PAME], [CODIGO EMPLEADO], dbo.ProperCase(APELLIDOS) AS APELLIDOS, dbo.ProperCase(NOMBRES) AS NOMBRES, [CARGO MILITAR], [CARGO PAME], [SALDO VACACIONAL ACUMULADO], ACREDITA, [VACACIONES CANCELADAS], [OMISIONES DE MARCAS], [AUSENCIA INJUSTIFICADA], AMP, AMN, SALDO FROM [VISTA CONSULTA GENERAL DE EXPEDIENTES SALDOS VAC] WHERE " & CADENA_NIVELES & " and " & SELECCION & ""
        'End If
        Call CARGAR_DATAGRIDVIEW()
        For I = 0 To Me.DGV1.ColumnCount - 1
            Me.DGV1.Columns(I).Width = 70
            Me.DGV1.Columns(I).Visible = True
            Me.DGV1.Columns(I).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            Me.DGV1.Columns(I).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Next
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call VERIFICAR_ACCESOS("304") : If HAY_ACCESO = False Then : Exit Sub : End If
        If Me.RadioButton1.Checked = False And Me.RadioButton2.Checked = False Then
            MsgBox("Debe seleccionar el Tipo de Estructura primeramente", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Try
            TITULO_EXCEL = "CONSULTA GENERAL CON SALDOS VACACIONALES"
            If Me.RadioButton1.Checked = True Then
                EXPORTAR_DATOS_A_EXCEL(DGV1, "Consulta General con Saldos Vacacionales [Estructura Militar] (" & UCase(Now) & ")")
            End If
            If Me.RadioButton2.Checked = True Then
                EXPORTAR_DATOS_A_EXCEL(DGV1, "Consulta General con Saldos Vacacionales [Estructura Pame] (" & UCase(Now) & ")")
            End If
        Catch ex As Exception
            MessageBox.Show("No se puede generar el documento Excel.", "Mensaje del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class