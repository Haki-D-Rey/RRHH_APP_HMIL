Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Public Class Frm_Expedientes_Inactivos
    Private Sub Frm_Expedientes_Inactivos_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Frm_Principal.Timer1.Stop()
        EXPEDIENTES_I = True
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
        Frm_Expedientes_Inactivos_Buscar.ShowDialog()
        If CERRAR = False Then
            tp = TabGeneral.TabPages(1)
            TabGeneral.SelectedTab = tp
            Me.Label4.Text = Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Text
            Call NO_ACTIVAR()
            Call SELECCIONAR_CADENA()
            Call DESACTIVAR_TODO()
            TabGeneral.Focus()
        End If
        '-------------------------------------------------
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        EXPEDIENTES_I = False
        Me.TextBox35.Text = ""
        Me.TextBox36.Text = ""
        vID_M_P = 0 : vID_M_C = 0
        Frm_Principal.Timer1.Start()
        Me.Close()
    End Sub

    Private Sub Frm_Expedientes_Inactivos_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            EXPEDIENTES_I = False
            Me.TextBox35.Text = ""
            Me.TextBox36.Text = ""
            vID_M_P = 0 : vID_M_C = 0
            Me.Close()
        End If
    End Sub
    Private Sub NO_ACTIVAR()
        Me.CheckBox5.Enabled = False
        Me.CheckBox11.Enabled = False
        Me.CheckBox6.Enabled = False
        Me.CheckBox7.Enabled = False
        Me.Button2.Enabled = False
        Me.DateTimePicker1.Enabled = False
        Me.DateTimePicker2.Enabled = False
        Me.DateTimePicker3.Enabled = False
        Me.Button2.Enabled = False  'EXAMINAR
    End Sub
    'Private Sub TabGeneral_DrawItem(sender As Object, e As DrawItemEventArgs) Handles TabGeneral.DrawItem
    '    Dim G As Graphics
    '    Dim STEXT As String
    '    Dim iX As Integer
    '    Dim iY As Integer
    '    Dim SIZETEXT As SizeF
    '    Dim CTLTAB As TabControl

    '    CTLTAB = CType(sender, TabControl)
    '    G = e.Graphics

    '    STEXT = CTLTAB.TabPages(e.Index).Text
    '    SIZETEXT = G.MeasureString(STEXT, CTLTAB.Font)
    '    iX = e.Bounds.Left + 6
    '    iY = e.Bounds.Top + (e.Bounds.Height - SIZETEXT.Height) / 2
    '    G.DrawString(STEXT, CTLTAB.Font, Brushes.Black, iX, iY)
    'End Sub
    Private Sub TabGeneral_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TabGeneral.Selecting
        Call SELECCIONAR_CADENA()
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

    End Sub
    Private Sub LIMPIAR_TEXTOS_2()
        Me.TextBox5.Text = "" : Me.TextBox6.Text = "" : Me.TextBox30.Text = "" : Me.TextBox9.Text = ""
        Me.TextBox10.Text = "" : Me.TextBox11.Text = "" : Me.TextBox12.Text = "" : Me.TextBox13.Text = ""
        Me.TextBox7.Text = "" : Me.TextBox8.Text = "" : Me.TextBox14.Text = "" : Me.TextBox15.Text = ""
        Me.TextBox16.Text = "" : Me.TextBox17.Text = "" : Me.TextBox18.Text = "" : Me.TextBox21.Text = ""
        Me.TextBox20.Text = "" : Me.TextBox19.Text = "" : Me.TextBox22.Text = "" : Me.TextBox23.Text = ""
        Me.TextBox24.Text = "" : Me.TextBox25.Text = "" : Me.TextBox26.Text = "" : Me.TextBox27.Text = ""
        Me.TextBox28.Text = ""
        Me.TextBox1.Text = "" : Me.TextBox2.Text = "" : Me.TextBox3.Text = ""
        Me.TextBox42.Text = ""
        Me.TextBox34.Text = "" : Me.PictureBox4.Image = Nothing
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub
    Private Sub SELECCIONAR_CADENA()
        Me.Label4.Text = Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Text
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 1 Then     'INFORMACION ESTRUCTURA
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
            If isuID_TIPO_USUARIO = 5 Then
                CADENAsql = "SELECT * FROM [VISTA MAESTRO DE PERSONAS] WHERE ID_M_P = " & vID_M_P & "  AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ID_M_P"
            Else
                CADENAsql = "SELECT * FROM [VISTA MAESTRO DE PERSONAS] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P"
            End If
            Call BUSCAR_VISTA_MAESTRO_DE_PERSONAS()
            Call DESACTIVAR_TODO()
            Me.TextBox35.Text = vID_M_P
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 3 Then     'NÚCLEO FAMILIAR
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE NUCLEO FAMILIAR] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_NF"
            Call CARGAR_DATAGRIDVIEW_1()
            Call ARMAR_DATAGRIDVIEW_1()
            Me.TextBox35.Text = vID_M_P
            Call DESACTIVAR_TODO()
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 4 Then     'EXPERIENCIA LABORAL
            CADENAsql = "SELECT * FROM [MAESTRO DE EXPERIENCIA LABORAL] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_EL"
            Call CARGAR_DATAGRIDVIEW_2()
            Call ARMAR_DATAGRIDVIEW_2()
            Me.TextBox35.Text = vID_M_P
            Call DESACTIVAR_TODO()
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 5 Then     'REFERENCIAS PERSONALES
            CADENAsql = "SELECT * FROM [MAESTRO DE REFERENCIAS PERSONALES] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_RP"
            Call CARGAR_DATAGRIDVIEW_3()
            Call ARMAR_DATAGRIDVIEW_3()
            Me.TextBox35.Text = vID_M_P
            Call DESACTIVAR_TODO()
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 6 Then     'ESTUDIOS
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE ESTUDIOS] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_M_EST"
            Call CARGAR_DATAGRIDVIEW_4()
            Call ARMAR_DATAGRIDVIEW_4()
            Me.TextBox35.Text = vID_M_P
            Call DESACTIVAR_TODO()
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 7 Then     'HABILIDADES ADMINISTRATIVAS
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HABILIDADES ADMIN] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_H_ADMIN"
            Call CARGAR_DATAGRIDVIEW_5()
            Call ARMAR_DATAGRIDVIEW_5()
            Me.TextBox35.Text = vID_M_P
            Call DESACTIVAR_TODO()
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 8 Then     'IDIOMAS
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE IDIOMAS] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_M_I"
            Call CARGAR_DATAGRIDVIEW_6()
            Call ARMAR_DATAGRIDVIEW_6()
            Me.TextBox35.Text = vID_M_P
            Call DESACTIVAR_TODO()
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 9 Then     'CURVA Y TALLA
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE CURVA Y TALLA] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_CT"
            Call CARGAR_DATAGRIDVIEW_7()
            Call ARMAR_DATAGRIDVIEW_7()
            Me.TextBox35.Text = vID_M_P
            Call DESACTIVAR_TODO()
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 10 Then    'CLASIFICACION DE PERSONAL
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE CLASIFICACION DE PERSONAL] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_C_PERSONAL"
            Call CARGAR_DATAGRIDVIEW_8()
            Call ARMAR_DATAGRIDVIEW_8()
            Me.TextBox35.Text = vID_M_P
            Call DESACTIVAR_TODO()
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 11 Then    'CONDECORACIONES
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE CONDECORACIONES] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_M_CONDE"
            Call CARGAR_DATAGRIDVIEW_9()
            Call ARMAR_DATAGRIDVIEW_9()
            Me.TextBox35.Text = vID_M_P
            Call DESACTIVAR_TODO()
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 12 Then    'EVENTUALIDADES
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE EVENTUALIDADES] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_M_EV"
            Call CARGAR_DATAGRIDVIEW_10()
            Call ARMAR_DATAGRIDVIEW_10()
            Me.TextBox35.Text = vID_M_P
            Call DESACTIVAR_TODO()
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 13 Then    'NOMINAS Y SALARIOS
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE NOMINAS Y SALARIOS] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_NOMINA, F_MOVIMIENTO_N DESC"
            Call CARGAR_DATAGRIDVIEW_11()
            Call ARMAR_DATAGRIDVIEW_11()
            Me.TextBox35.Text = vID_M_P
            Call DESACTIVAR_TODO()
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 14 Then    'DOCUMENTOS DIGITALES
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
            CADENAsql = "SELECT * FROM [VISTA HISTORICO DE MOVIMIENTO] WHERE ID_M_P = " & vID_M_P & " AND APLICAR = 1 ORDER BY FEC_MOV DESC"
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
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HORARIOS] WHERE ID_M_P = " & vID_M_P & " ORDER BY HORA_E"
            Call CARGAR_DATAGRIDVIEW_16()
            Call ARMAR_DATAGRIDVIEW_16()
            Me.TextBox35.Text = vID_M_P
            Call DESACTIVAR_TODO()
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 19 Then    'SUBSIDIOS
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE SUBSIDIOS] WHERE ID_M_P = " & vID_M_P & " ORDER BY FECHA_I"
            Call CARGAR_DATAGRIDVIEW_17()
            Call ARMAR_DATAGRIDVIEW_17()
            Me.TextBox35.Text = vID_M_P
            Call DESACTIVAR_TODO()
            GoTo SALIR
        End If
        If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 20 Then     'CONTRATOS Y EVALUACIONES
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
    Private Sub CARGAR_DATAGRIDVIEW_19()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE INMUNIZACIONES]")
        DGV19.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE INMUNIZACIONES]")
        Me.Label7.Text = DGV19.RowCount & " Registro(s)"
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
        Me.Label109.Text = "Total de Registros: " & DGV18.RowCount
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
        Me.Label103.Text = "Total de Registros: " & DGV17.RowCount
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
        Me.Label102.Text = "Total de Registros: " & DGV16.RowCount
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
        Me.Label97.Text = "Total de Registros: " & DGV15.RowCount
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
        Me.Label96.Text = "Total de Registros: " & DGV14.RowCount
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
    Private Sub CARGAR_DATAGRIDVIEW_12()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE DOCUMENTOS DIGITALES]")
        DGV12.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE DOCUMENTOS DIGITALES]")
        Me.Label90.Text = "Total de Registros: " & DGV12.RowCount
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
        Me.Label89.Text = "Total de Registros: " & DGV11.RowCount
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
        Me.Label88.Text = "Total de Registros: " & DGV10.RowCount
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
        Me.Label87.Text = "Total de Registros: " & DGV9.RowCount
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
        Me.Label86.Text = "Total de Registros: " & DGV8.RowCount
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
        Me.Label85.Text = "Total de Registros: " & DGV7.RowCount
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
        Me.Label84.Text = "Total de Registros: " & DGV6.RowCount
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

        Me.DGV6.Columns(6).HeaderText = "Lee"
        Me.DGV6.Columns(6).Width = 70
        Me.DGV6.Columns(6).Visible = True
        Me.DGV6.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV6.Columns(7).HeaderText = "%"
        Me.DGV6.Columns(7).Width = 70
        Me.DGV6.Columns(7).Visible = True
        Me.DGV6.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV6.Columns(8).HeaderText = "Escribe"
        Me.DGV6.Columns(8).Width = 70
        Me.DGV6.Columns(8).Visible = True
        Me.DGV6.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV6.Columns(9).HeaderText = "%"
        Me.DGV6.Columns(9).Width = 70
        Me.DGV6.Columns(9).Visible = True
        Me.DGV6.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

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
        Me.Label83.Text = "Total de Registros: " & DGV5.RowCount
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
        Me.Label82.Text = "Total de Registros: " & DGV4.RowCount
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
        Me.Label81.Text = "Total de Registros: " & DGV3.RowCount
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
        Me.Label80.Text = "Total de Registros: " & DGV2.RowCount
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
        Me.Label61.Text = "Total de Registros: " & DGV1.RowCount
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
                Me.TextBox26.Text = MiDataTable.Rows(0).Item(53).ToString
                Me.TextBox27.Text = MiDataTable.Rows(0).Item(54).ToString
                Me.ComboBox23.Text = MiDataTable.Rows(0).Item(56).ToString
                Me.ComboBox22.Text = MiDataTable.Rows(0).Item(58).ToString
                'Me.TextBox29.Text = MiDataTable.Rows(0).Item(59).ToString
                Me.TextBox28.Text = MiDataTable.Rows(0).Item(60).ToString
                Me.CheckBox5.Checked = MiDataTable.Rows(0).Item(63).ToString
                Me.CheckBox11.Checked = MiDataTable.Rows(0).Item(92).ToString()
                'FOTO
                Call CARGAR_FOTO()
                '---
                Me.TextBox1.Text = MiDataTable.Rows(0).Item(7).ToString
                Me.TextBox2.Text = MiDataTable.Rows(0).Item(4).ToString
                Me.TextBox3.Text = Mid(MiDataTable.Rows(0).Item(2).ToString, 1, 10)
                '---
                Me.ComboBox24.Text = MiDataTable.Rows(0).Item(68).ToString
                Me.ComboBox25.Text = MiDataTable.Rows(0).Item(70).ToString
                'Me.TextBox43.Text = MiDataTable.Rows(0).Item(79).ToString
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
        Call VERIFICAR_ACCESOS("158") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Expedientes_Inactivos_Buscar.ShowDialog()
        If CERRAR = False Then
            Dim tp As TabPage = TabGeneral.TabPages(1)
            TabGeneral.SelectedTab = tp
            Me.Label4.Text = Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Text
            Call NO_ACTIVAR()
            Call SELECCIONAR_CADENA()
            Call DESACTIVAR_TODO()
            TabGeneral.Focus()
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs)
        'Frm_Expedientes_Activos_Buscar.ShowDialog()
    End Sub
    Private Sub ACTIVAR_01()
        '-------------------------------------
        'BLOQUEAR CONTROLES TAB : 0
        BLOQUEO = 0
        Call BLOQUEAR_OBJETOS()
        '-------------------------------------
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
        Me.TextBox34.ReadOnly = False : Me.CheckBox5.Enabled = True
        Me.CheckBox6.Enabled = True
        Me.CheckBox7.Enabled = True
        Me.CheckBox8.Enabled = True
        Me.Button2.Enabled = True   'EXAMINAR
        '-------------------------------------
        'BLOQUEAR CONTROLES TAB : 1
        BLOQUEO = 1
        Call BLOQUEAR_OBJETOS()
        '-------------------------------------
        Me.DateTimePicker1.Focus()
    End Sub
    Private Sub DESACTIVAR_TODO()
        '0
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
        Me.TextBox34.ReadOnly = True : Me.CheckBox5.Enabled = False
        Me.CheckBox6.Enabled = False
        Me.CheckBox7.Enabled = False
        Me.CheckBox8.Enabled = False
        Me.CheckBox9.Enabled = False
        Me.CheckBox10.Enabled = False
        Me.Button2.Enabled = False  'EXAMINAR
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
            Me.DateTimePicker4.Enabled = True
            Me.DateTimePicker5.Enabled = True
        Else
            Me.DateTimePicker4.Enabled = False
            Me.DateTimePicker5.Enabled = False
        End If
    End Sub
    Dim BLOQUEO As Integer
    Private Sub BLOQUEAR_OBJETOS()
        If BLOQUEO = 0 Then : Me.TabGeneral.TabPages(0).Enabled = True : Else : Me.TabGeneral.TabPages(0).Enabled = False : End If
        If BLOQUEO = 1 Then : Me.TabGeneral.TabPages(1).Enabled = True : Else : Me.TabGeneral.TabPages(1).Enabled = False : End If
        If BLOQUEO = 2 Then : Me.TabGeneral.TabPages(2).Enabled = True : Else : Me.TabGeneral.TabPages(2).Enabled = False : End If
        If BLOQUEO = 3 Then : Me.TabGeneral.TabPages(3).Enabled = True : Else : Me.TabGeneral.TabPages(3).Enabled = False : End If
        If BLOQUEO = 4 Then : Me.TabGeneral.TabPages(4).Enabled = True : Else : Me.TabGeneral.TabPages(4).Enabled = False : End If
        If BLOQUEO = 5 Then : Me.TabGeneral.TabPages(5).Enabled = True : Else : Me.TabGeneral.TabPages(5).Enabled = False : End If
        If BLOQUEO = 6 Then : Me.TabGeneral.TabPages(6).Enabled = True : Else : Me.TabGeneral.TabPages(6).Enabled = False : End If
        If BLOQUEO = 7 Then : Me.TabGeneral.TabPages(7).Enabled = True : Else : Me.TabGeneral.TabPages(7).Enabled = False : End If
        If BLOQUEO = 8 Then : Me.TabGeneral.TabPages(8).Enabled = True : Else : Me.TabGeneral.TabPages(8).Enabled = False : End If
        If BLOQUEO = 9 Then : Me.TabGeneral.TabPages(9).Enabled = True : Else : Me.TabGeneral.TabPages(9).Enabled = False : End If
        If BLOQUEO = 10 Then : Me.TabGeneral.TabPages(10).Enabled = True : Else : Me.TabGeneral.TabPages(10).Enabled = False : End If
        If BLOQUEO = 11 Then : Me.TabGeneral.TabPages(11).Enabled = True : Else : Me.TabGeneral.TabPages(11).Enabled = False : End If
        If BLOQUEO = 12 Then : Me.TabGeneral.TabPages(12).Enabled = True : Else : Me.TabGeneral.TabPages(12).Enabled = False : End If
        If BLOQUEO = 13 Then : Me.TabGeneral.TabPages(13).Enabled = True : Else : Me.TabGeneral.TabPages(13).Enabled = False : End If
        If BLOQUEO = 14 Then : Me.TabGeneral.TabPages(14).Enabled = True : Else : Me.TabGeneral.TabPages(14).Enabled = False : End If
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs)
        Dim tp As TabPage = TabGeneral.TabPages(1)
        TabGeneral.SelectedTab = tp
        Me.Label4.Text = Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Text
        Call NO_ACTIVAR()
        Call SELECCIONAR_CADENA()
        TabGeneral.Focus()
        Call DESACTIVAR_TODO()
    End Sub
    Private Sub Button48_Click(sender As Object, e As EventArgs)
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
    Private Sub TextBox26_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox26.KeyDown
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
    Private Sub TextBox28_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox28.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox29_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox33_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox32_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox31_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    'VISTA MAESTRO DE DOCUMENTOS DIGITALES
    Private Sub DGV12_Click(sender As Object, e As EventArgs) Handles DGV12.Click
        If Me.DGV12.RowCount > 0 Then
            For I = 0 To Me.DGV12.RowCount - 1
                If DGV12.Rows(I).Selected = True Then
                    vmdID_M_DD = Me.DGV12.Item(0, I).Value
                    vmdID_M_P = Me.DGV12.Item(1, I).Value
                    vmdID_D1 = Me.DGV12.Item(2, I).Value
                    vmdDIRECTORIO = Me.DGV12.Item(3, I).Value
                    vmdID_T_DOC = Me.DGV12.Item(4, I).Value
                    vmdD_T_DOC = Me.DGV12.Item(5, I).Value
                    vmdNOMBRE = Me.DGV12.Item(6, I).Value
                    vmdOBSERVACION = Me.DGV12.Item(7, I).Value
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub Button42_Click(sender As Object, e As EventArgs) Handles Button42.Click   'ELIMINAR
        Call VERIFICAR_ACCESOS("178") : If HAY_ACCESO = False Then : Exit Sub : End If
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
        Call VERIFICAR_ACCESOS("177") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmdID_M_DD = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Editar_Inactivos_Documentos.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE DOCUMENTOS DIGITALES] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P, ID_T_DOC"
            Call CARGAR_DATAGRIDVIEW_12()
            Call ARMAR_DATAGRIDVIEW_12()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_12()
        End If
    End Sub
    Private Sub Button43_Click(sender As Object, e As EventArgs) Handles Button43.Click 'EDITAR
        Call VERIFICAR_ACCESOS("177") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmdID_M_DD = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Editar_Inactivos_Documentos.ShowDialog()
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
        Call VERIFICAR_ACCESOS("176") : If HAY_ACCESO = False Then : Exit Sub : End If
        ''-------------------------------
        Frm_Expedientes_Inactivos_Agregar.ShowDialog()
    End Sub
    Private Sub Button51_Click(sender As Object, e As EventArgs) Handles Button51.Click 'VISUALIZAR
        Call VERIFICAR_ACCESOS("175") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vmdID_M_DD = 0 Then
            MsgBox("Debe seleccionar el registro que desea Visualizar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Visor_Documentos_Inactivos.ShowDialog()
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
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Frm_Expedientes_Inactivos_Imprimir.ShowDialog()
    End Sub
    Private Sub DGV15_RowPrePaint(sender As Object, e As DataGridViewRowPrePaintEventArgs) Handles DGV15.RowPrePaint
        Call VERIFICAR_ACCESOS("159") : If HAY_ACCESO = False Then : Exit Sub : End If
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
End Class