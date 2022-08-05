Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Public Class Frm_Movimientos_Exp_Add_Recurso_Nuevo
    Private Sub Frm_Movimientos_Exp_Add_Recurso_Nuevo_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Movimientos_Exp_Add_Recurso_Nuevo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call Clases.BUSCAR_FECHA_Y_HORA()
        Me.DateTimePicker1.Value = FECHA_SERVIDOR
        Me.DateTimePicker2.Value = Frm_Movimientos_Exp_Add.DateTimePicker1.Value
        Me.DateTimePicker3.Value = FECHA_SERVIDOR : Me.DateTimePicker3.Checked = False
        Me.TextBox5.Text = ""
        Me.TextBox6.Text = ""
        Me.TextBox30.Text = ""
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_SEXO()
        Me.TextBox9.Text = ""
        Call CARGAR_COMBO_CAT_DE_IPSS()
        Me.TextBox10.Text = ""
        Me.TextBox12.Text = ""
        Me.TextBox13.Text = ""
        Me.TextBox11.Text = ""
        Call CARGAR_COMBO_CAT_DE_CATEGORIA_DE_LIC()
        Call CARGAR_COMBO_CAT_DE_TIPO_SANGUINEO()
        Me.TextBox7.Text = ""
        Me.TextBox8.Text = ""
        Me.TextBox14.Text = ""
        Me.TextBox15.Text = ""
        Me.TextBox16.Text = ""
        Me.TextBox17.Text = ""
        Me.TextBox18.Text = ""
        Me.TextBox21.Text = ""
        Me.TextBox20.Text = ""
        Me.TextBox19.Text = ""
        Me.TextBox22.Text = ""
        Me.TextBox23.Text = ""
        Call CARGAR_COMBO_CAT_DE_TEZ()
        Call CARGAR_COMBO_CAT_DE_CABELLO()
        Call CARGAR_COMBO_CAT_DE_NIVEL_ACADEMICO()
        Call CARGAR_COMBO_CAT_DE_NIVEL_PROFESIONAL()
        Call CARGAR_COMBO_CAT_DE_NACIONALIDAD_N()
        Call CARGAR_COMBO_CAT_DE_DPTO_N()
        Call CARGAR_COMBO_CAT_DE_MUNICIPIO_N()
        Me.TextBox24.Text = ""
        Call CARGAR_COMBO_CAT_DE_NACIONALIDAD_D()
        Call CARGAR_COMBO_CAT_DE_DPTO_D()
        Call CARGAR_COMBO_CAT_DE_MUNICIPIO_D()
        Me.TextBox25.Text = ""
        Call CARGAR_COMBO_CAT_DE_ESTADO_CIVIL()
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_HORARIO()
        Me.TextBox26.Text = ""
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_AFILIACION()
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_PLAZA()
        Call CARGAR_COMBO_CAT_DE_FUNCIONALIDAD()
        Me.TextBox27.Text = ""
        Call CARGAR_COMBO_CAT_DE_NIVEL_DE_COMPETENCIA()
        Me.TextBox44.Text = ""
        Me.TextBox34.Text = ""
        Me.TextBox2.Text = ""
        Call CARGAR_COMBO_CAT_DE_ADMINISTRADORES()
        Me.CheckBox5.Checked = False
        Me.CheckBox11.Checked = False
        Me.CheckBox8.Checked = False
        Me.CheckBox6.Checked = False
        Me.CheckBox7.Checked = False
        Me.CheckBox9.Checked = True
        Me.CheckBox10.Checked = False
        Me.PictureBox4.Image = Me.PictureBox7.Image
        '-------------------------------------
        Call GENERAR_CODIGO_EMPLEADO()
        TextBox5.Text = Format(Val(CODIGO_EMPLEADO), "0000000000")
        '-------------------------------------
        Me.DateTimePicker1.Focus()
    End Sub
    Dim CODIGO_EMPLEADO As Integer
    Private Sub GENERAR_CODIGO_EMPLEADO()
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
    Private Sub CARGAR_COMBO_CAT_DE_NIVEL_DE_COMPETENCIA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE NIVEL DE COMPETENCIA] ORDER BY DESCRIPCION", Conexion.CxRRHH)
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
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE FUNCIONALIDAD] ORDER BY DESCRIPCION", Conexion.CxRRHH)
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
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_PLAZA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE PLAZA] ORDER BY DESCRIPCION", Conexion.CxRRHH)
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
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_AFILIACION()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE AFILIACION] ORDER BY DESCRIPCION", Conexion.CxRRHH)
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
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_HORARIO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE HORARIO] ORDER BY DESCRIPCION", Conexion.CxRRHH)
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
    Private Sub CARGAR_COMBO_CAT_DE_ESTADO_CIVIL()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTADO CIVIL] ORDER BY DESCRIPCION", Conexion.CxRRHH)
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
    Private Sub CARGAR_COMBO_CAT_DE_MUNICIPIO_D()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE MUNICIPIO_D] ORDER BY DESCRIPCION", Conexion.CxRRHH)
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
    Private Sub CARGAR_COMBO_CAT_DE_DPTO_D()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE DPTO_D] ORDER BY DESCRIPCION", Conexion.CxRRHH)
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
    Private Sub CARGAR_COMBO_CAT_DE_NACIONALIDAD_D()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE NACIONALIDAD_D] ORDER BY DESCRIPCION", Conexion.CxRRHH)
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
    Private Sub CARGAR_COMBO_CAT_DE_MUNICIPIO_N()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE MUNICIPIO_N] ORDER BY DESCRIPCION", Conexion.CxRRHH)
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
    Private Sub CARGAR_COMBO_CAT_DE_DPTO_N()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE DPTO_N] ORDER BY DESCRIPCION", Conexion.CxRRHH)
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
    Private Sub CARGAR_COMBO_CAT_DE_NACIONALIDAD_N()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE NACIONALIDAD_N] ORDER BY DESCRIPCION", Conexion.CxRRHH)
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
    Private Sub CARGAR_COMBO_CAT_DE_NIVEL_PROFESIONAL()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE NIVEL PROFESIONAL] ORDER BY DESCRIPCION", Conexion.CxRRHH)
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
    Private Sub CARGAR_COMBO_CAT_DE_NIVEL_ACADEMICO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE NIVEL ACADEMICO] ORDER BY DESCRIPCION", Conexion.CxRRHH)
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
    Private Sub CARGAR_COMBO_CAT_DE_CABELLO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE CABELLO] ORDER BY DESCRIPCION", Conexion.CxRRHH)
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
    Private Sub CARGAR_COMBO_CAT_DE_TEZ()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TEZ] ORDER BY DESCRIPCION", Conexion.CxRRHH)
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
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_SANGUINEO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO SANGUINEO] ORDER BY DESCRIPCION", Conexion.CxRRHH)
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
    Private Sub CARGAR_COMBO_CAT_DE_CATEGORIA_DE_LIC()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE CATEGORIA DE LIC] ORDER BY DESCRIPCION", Conexion.CxRRHH)
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
    Private Sub CARGAR_COMBO_CAT_DE_IPSS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE IPSS] ORDER BY DESCRIPCION", Conexion.CxRRHH)
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
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE SEXO] ORDER BY DESCRIPCION", Conexion.CxRRHH)
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
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        CERRAR = True
        Me.Close()
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
    Private Sub ComboBox26_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox26.KeyDown
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
    Private Sub TextBox15_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox15.KeyDown
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
    Private Sub TextBox17_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox17.KeyDown
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
    Private Sub TextBox20_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox20.KeyDown
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
    Private Sub TextBox22_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox22.KeyDown
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
    Private Sub ComboBox12_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox12.KeyDown
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
    Private Sub ComboBox19_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox19.KeyDown
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
    Private Sub TextBox25_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox25.KeyDown
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
    Private Sub ComboBox23_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox23.KeyDown
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
    Private Sub ComboBox27_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox27.KeyDown
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
    Private Sub ComboBox28_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox28.KeyDown
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
    Private Sub TextBox28_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox43_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox44_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox44.KeyDown
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
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
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
    Private Sub CheckBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox8.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub CheckBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox6.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub CheckBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox7.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub CheckBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox9.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub CheckBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox10.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Dim CONTROL_FOTO As Boolean
    Dim data() As Byte
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim ios As FileStream
        Try
            Me.OpenFileDialog1.Filter = "Imagenes (JPG)|*.jpg"
            If Me.OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.Cancel Then
                Exit Sub
            Else
                ios = New FileStream(Me.OpenFileDialog1.FileName, FileMode.Open, FileAccess.Read)
                ReDim data(ios.Length)
                ios.Read(data, 0, CInt(ios.Length))
                Me.PictureBox4.SizeMode = PictureBoxSizeMode.StretchImage
                Me.PictureBox4.Load(Me.OpenFileDialog1.FileName)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Mensaje del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification) 'en caso de error muestre un mensaje
        End Try
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
            MiDataAdapter = New SqlDataAdapter("SELECT ID_M_P FROM [MAESTRO DE PERSONAS] ORDER BY ID_M_P", Conexion.CxRRHH)
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
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [MAESTRO DE PERSONAS] WHERE CODIGO = '" & Me.TextBox5.Text & "' ORDER BY CODIGO", Conexion.CxRRHH)
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
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If Me.DateTimePicker2.Value <= Me.DateTimePicker1.Value Then
            MsgBox("La Fecha de Ingreso al PAME no puede ser menor o igual a la Fecha de Nacimiento del Recurso", vbInformation, "Mensaje del Sistema")
            Me.DateTimePicker1.Focus()
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
        If Len(Me.TextBox6.Text) < 3 Then
            MsgBox("Debe Digitar Correctamente los Apellidos del Recurso que desea agregar", vbInformation, "Mensaje del Sistema")
            Me.TextBox6.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox30.Text) < 3 Then
            MsgBox("Debe Digitar Correctamente los Nombres del Recurso que desea agregar", vbInformation, "Mensaje del Sistema")
            Me.TextBox30.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox9.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Tipo de Sexo Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox9.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox10.Text) <> 14 Then
            MsgBox("Debe Digitar la Cedula del Recurso Correctamente, recuerde que debe tener 14 caracteres", vbInformation, "Mensaje del Sistema")
            Me.TextBox10.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox26.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar la IPSS Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox26.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox24.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar la Categoria de Licencia de Conducir Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox24.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox25.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Tipo Sanguineo Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox25.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox7.Text) < 5 Then
            MsgBox("Debe Digitar la Dirección Domiciliar 1 del Recurso Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox7.Focus()
            Exit Sub
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
        If Len(Me.ComboBox13.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Nivel Academico Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox13.Focus()
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
            MsgBox("Debe seleccionar de Departamento de Nacimiento Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox14.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox16.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Municipio de Nacimiento Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox16.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox19.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar la Nacionalidad Domiciliar Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox19.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox18.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar de Departamento Domiciliar Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox18.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox17.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Municipio Domiciliar Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox17.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox25.Text) < 5 Then
            MsgBox("Debe Digitar el Lugar de Ubicacion o Referencia Domiciliar más próxima de Recurso Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox25.Focus()
            Exit Sub
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
        If Len(Me.ComboBox22.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Tipo de Plaza Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox22.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox23.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Tipo de Afiliacion Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox23.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox27.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Tipo de Funcionalidad Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox27.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox28.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Nivel de Commpetencia Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox28.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox44.Text) < 10 Then
            MsgBox("Debe Digitar el Apellido y Nombre del Jefe Inmediato del Recurso Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox44.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox2.Text) < 3 Then
            MsgBox("Debe Digitar el Prefijo del Recurso a Utilizar Correctamente", vbInformation, "Mensaje del Sistema")
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

        Call GRABAR_REGISTRO_DE_PERSONA()

        cFECHAING = Mid(Me.DateTimePicker2.Value, 1, 10)
        cEMPLEADO = Trim(Me.TextBox5.Text)
        bAN = Trim(Me.TextBox6.Text) & " " & Trim(Me.TextBox30.Text)
        vID_M_P = Me.TextBox1.Text
        CERRAR = False
        Me.Close()
    End Sub
    Dim cadena_CAMPOS1, cadena_VALORES1 As String
    Private Sub Button47_Click(sender As Object, e As EventArgs) Handles Button47.Click
        Frm_Expedientes_Activos_Jefe_Inmediato.ShowDialog()
        If CERRAR = False Then
            Me.TextBox44.Text = bEMPLEADO
        End If
    End Sub
    Private Sub Button63_Click(sender As Object, e As EventArgs) Handles Button63.Click
        Frm_Expedientes_Activos_Prefijos.ShowDialog()
        If CERRAR = False Then
            If bPREFIJO <> "" Then
                Me.TextBox2.Text = bPREFIJO
            End If
        End If
    End Sub
    Private Sub GRABAR_REGISTRO_DE_PERSONA()

        cadena_CAMPOS1 = "ID_M_P, FEC_NAC, FEC_ING_PAME, FEC_ING_EN, " &
                "CODIGO, APELLIDOS, NOMBRES, ID_T_SEXO, DIRECCION_1, " &
                "DIRECCION_2, NSS, N_CEDULA, N_LIC, N_MINSA, " &
                "LIC_IMAGEN, TELEFONO_1, TELEFONO_2, CONVENCIONAL, CORREO_1, " &
                "CORREO_2, WHATSAPP, FACEBOOK, TWITTER, PESO, " &
                "ESTATURA, ID_TEZ, ID_CABELLO, ID_N_ACADEMICO, ID_N_PROFESIONAL, " &
                "ID_NACIONALIDAD_N, ID_DPTO_N, ID_M_N, LUGAR_N, ID_NACIONALIDAD_D, " &
                "ID_DPTO_D, ID_M_D, LUGAR_D, ID_E_CIVIL, ID_T_HORARIO, " &
                "N_C_BANCO, EMPLEADO_ANT, ID_T_AFILIACION, ID_T_PLAZA, D_N1, " &
                "D_N2, FOTO, ID_ESTADO_P, BECADO, ID_C_LIC, " &
                "ID_T_SANGUINEO, OBSERVACIONES, ID_IPSS, ESPECIALISTA, SUBESPECIALISTA, " &
                "ID_F, EN_SITUACION, CARGO_REAL, JEFE_REAL, CALCULO_VACACIONAL, " &
                "ID_N_COMPETENCIA, SERVICIOS_PROFESIONALES, PREFIJO, ID_ADMIN, PRACTICANTE"

        Dim ms As New System.IO.MemoryStream() 'creando el memoryStream
        Me.PictureBox4.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)

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

        cadena_VALORES1 = "" & Me.TextBox1.Text & ", " & fechaNacimiento & ", " & fechaingreso & ", " & fechaingresoEN & ", " &
                 "'" & Me.TextBox5.Text & "', '" & Me.TextBox6.Text & "', '" & Me.TextBox30.Text & "', " & Me.ComboBox9.SelectedValue & ", '" & Me.TextBox7.Text & "', " &
                 "'" & Me.TextBox8.Text & "', '" & Me.TextBox9.Text & "', '" & Me.TextBox10.Text & "', '" & Me.TextBox11.Text & "', '" & Me.TextBox12.Text & "', " &
                 "'" & Me.TextBox13.Text & "', '" & Me.TextBox14.Text & "', '" & Me.TextBox15.Text & "', '" & Me.TextBox16.Text & "', '" & Me.TextBox17.Text & "', " &
                 "'" & Me.TextBox18.Text & "', '" & Me.TextBox21.Text & "', '" & Me.TextBox20.Text & "', '" & Me.TextBox19.Text & "', '" & Me.TextBox22.Text & "', " &
                 "'" & Me.TextBox23.Text & "', " & Me.ComboBox10.SelectedValue & ", " & Me.ComboBox11.SelectedValue & ", " & Me.ComboBox13.SelectedValue & ", " & Me.ComboBox12.SelectedValue & ", " &
                 "" & Me.ComboBox15.SelectedValue & ", " & Me.ComboBox14.SelectedValue & ", " & Me.ComboBox16.SelectedValue & ", '" & Me.TextBox24.Text & "', " & Me.ComboBox19.SelectedValue & ", " &
                 "" & Me.ComboBox18.SelectedValue & ", " & Me.ComboBox17.SelectedValue & ", '" & Me.TextBox25.Text & "', " & Me.ComboBox20.SelectedValue & ", " & Me.ComboBox21.SelectedValue & ", " &
                 "'" & Me.TextBox26.Text & "', '" & Me.TextBox27.Text & "', " & Me.ComboBox23.SelectedValue & ", " & Me.ComboBox22.SelectedValue & ", NULL, " &
                 "NULL,  @data, 1, '" & Me.CheckBox5.Checked & "', " & Me.ComboBox24.SelectedValue & ", " &
                 "" & Me.ComboBox25.SelectedValue & ", '" & Me.TextBox34.Text & "', " & Me.ComboBox26.SelectedValue & ", '" & Me.CheckBox6.Checked & "', '" & Me.CheckBox7.Checked & "', " &
                 "" & Me.ComboBox27.SelectedValue & ", '" & Me.CheckBox8.Checked & "', NULL, '" & Me.TextBox44.Text & "', '" & Me.CheckBox9.Checked & "', " &
                 "" & Me.ComboBox28.SelectedValue & ", '" & Me.CheckBox10.Checked & "', '" & Me.TextBox2.Text & "', " & Me.ComboBox33.SelectedValue & ", '" & Me.CheckBox11.Checked & "'"

        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim SQL As String
        Try
            SQL = "INSERT INTO [MAESTRO DE PERSONAS] (" & cadena_CAMPOS1 & ") values (" & cadena_VALORES1 & ")"
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
    End Sub
    Private Sub ComboBox33_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox33.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
End Class