Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Public Class Frm_Accidentes_Laborales
    Private Sub Frm_Accidentes_Laborales_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            HIGIENE_1 = False
            Me.TextBox36.Text = ""
            h1ID_M_P = 0
            Frm_Principal.Timer1.Start()
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Accidentes_Laborales_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Frm_Principal.Timer1.Stop()
        HIGIENE_1 = True
        Dim tp As TabPage = TabGeneral.TabPages(0)
        TabGeneral.SelectedTab = tp
        h1ID_M_P = 0
        Call DESACTIVAR_TODO()
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_SEXO()
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_HORARIO()
        Call CARGAR_COMBO_CAT_DE_ESTADO_DE_PERSONAS()
        If h1ID_M_P <> 0 Then
            Call SELECCIONAR_CADENA()
        End If
        Me.TextBox36.Text = ""
        TabGeneral.Focus()
        Me.Show()
        '-------------------------------------------------
        Frm_Accidentes_Laborales_Buscar.ShowDialog()
        If CERRAR = False Then
            tp = TabGeneral.TabPages(0)
            TabGeneral.SelectedTab = tp
            Call BUSCAR_PERSONA()
            Call BUSCAR_UBICACION_Y_CARGO()
            Call DESACTIVAR_TODO()
            If h1ID_M_P <> 0 Then
                'INFORMACION DEL ACCIDENTE
                CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HIGIENE INFORMACION DEL ACCIDENTE] WHERE ID_M_P = " & h1ID_M_P & " ORDER BY ID_M_P, NO_PARTE_ACC"
                Call CARGAR_DATAGRIDVIEW_1()
                Call ARMAR_DATAGRIDVIEW_1()
                Call SELECCIONAR_CADENA()
            End If
            TabGeneral.Focus()
        End If
        '-------------------------------------------------
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTADO_DE_PERSONAS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTADO DE PERSONAS] order by ID_ESTADO_P", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_ESTADO_P"
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
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE HORARIO] order by ID_T_HORARIO", Conexion.CxRRHH)
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
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_SEXO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE SEXO] order by ID_T_SEXO", Conexion.CxRRHH)
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
    Private Sub Button7_Click(sender As Object, e As System.EventArgs) Handles Button7.Click
        HIGIENE_1 = False
        Me.TextBox36.Text = ""
        h1ID_M_P = 0
        Frm_Principal.Timer1.Start()
        Me.Close()
    End Sub
    Private Sub DESACTIVAR_TODO()
        Me.Button11.Enabled = False
        Me.Button10.Enabled = False
        Me.Button9.Enabled = False
        DGV1.Enabled = False

        Me.Button14.Enabled = False
        Me.Button13.Enabled = False
        Me.Button12.Enabled = False
        DGV2.Enabled = False

        Me.Button17.Enabled = False
        Me.Button16.Enabled = False
        Me.Button15.Enabled = False
        DGV3.Enabled = False

        Me.Button8.Enabled = False
        Me.Button5.Enabled = False
        Me.Button4.Enabled = False
        DGV4.Enabled = False
    End Sub
    Private Sub TabGeneral_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TabGeneral.Selecting
        If h1ID_M_P <> 0 Then
            Call SELECCIONAR_CADENA()
        End If
    End Sub
    Private Sub SELECCIONAR_CADENA()
        Me.TextBox36.Text = h1ID_M_P
        CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HIGIENE DATOS DEL EVENTO] WHERE ID_M_P = " & h1ID_M_P & " AND ID_M_IT = " & h2ID_M_IT & " ORDER BY FECHA_ACC"
        Call CARGAR_DATAGRIDVIEW_2()
        Call ARMAR_DATAGRIDVIEW_2()
        If h1ID_M_P <> 0 Then
            'DATOS DEL EVENTO
            If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 2 Then
                CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HIGIENE DATOS DEL EVENTO] WHERE ID_M_P = " & h1ID_M_P & " AND ID_M_IT = " & h2ID_M_IT & " ORDER BY FECHA_ACC"
                Call CARGAR_DATAGRIDVIEW_2()
                Call ARMAR_DATAGRIDVIEW_2()
                Me.TextBox36.Text = h1ID_M_P
                GoTo SALIR
            End If
            'CAUSAS, FACTORES Y GRAVEDAD
            If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 3 Then
                CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HIGIENE CAUSAS FACTORES Y GRAVEDAD] WHERE ID_M_P = " & h1ID_M_P & " AND ID_M_IT = " & h2ID_M_IT & " ORDER BY ID_CFG"
                Call CARGAR_DATAGRIDVIEW_3()
                Call ARMAR_DATAGRIDVIEW_3()
                Me.TextBox36.Text = h1ID_M_P
                GoTo SALIR
            End If
            'DISPOSICION FINAL Y ANEXOS
            If Me.TabGeneral.TabPages(Me.TabGeneral.SelectedIndex).Tag = 4 Then
                CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HIGIENE DISPOSICION FINAL Y ANEXOS] WHERE ID_M_P = " & h1ID_M_P & " AND ID_M_IT = " & h2ID_M_IT & " ORDER BY ID_DFA"
                Call CARGAR_DATAGRIDVIEW_4()
                Call ARMAR_DATAGRIDVIEW_4()
                Me.TextBox36.Text = h1ID_M_P
                'Call DESACTIVAR_TODO()
                GoTo SALIR
            End If
        End If
SALIR:
    End Sub
    Private Sub Frm_Accidentes_Laborales_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        HIGIENE_1 = False
    End Sub
    Private Sub Frm_Accidentes_Laborales_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        HIGIENE_1 = False
    End Sub
    Private Sub Frm_Accidentes_Laborales_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        HIGIENE_1 = False
    End Sub
    Private Sub Frm_Accidentes_Laborales_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        HIGIENE_1 = False
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call VERIFICAR_ACCESOS("189") : If HAY_ACCESO = False Then : Exit Sub : End If
        '-------------------------------------------------
        Frm_Accidentes_Laborales_Buscar.ShowDialog()
        If CERRAR = False Then
            Dim tp As TabPage = TabGeneral.TabPages(0)
            TabGeneral.SelectedTab = tp
            Call BUSCAR_PERSONA()
            Call BUSCAR_UBICACION_Y_CARGO()
            Call DESACTIVAR_TODO()
            If h1ID_M_P <> 0 Then
                'INFORMACION DEL ACCIDENTE
                CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HIGIENE INFORMACION DEL ACCIDENTE] WHERE ID_M_P = " & h1ID_M_P & " ORDER BY ID_M_P, NO_PARTE_ACC"
                Call CARGAR_DATAGRIDVIEW_1()
                Call ARMAR_DATAGRIDVIEW_1()
                Call SELECCIONAR_CADENA()
            End If
            Call SELECCIONAR_CADENA()
            TabGeneral.Focus()
        End If
        '-------------------------------------------------
    End Sub
    Private Sub BUSCAR_PERSONA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT ID_M_P, CODIGO, APELLIDOS, NOMBRES, FEC_ING_PAME, NSS, N_CEDULA, EDAD, D_T_SEXO, D_T_HORARIO, D_ESTADO_P FROM [dbo].[VISTA MAESTRO DE PERSONAS HIGIENE] WHERE ID_M_P = '" & h1ID_M_P & "' ORDER BY ID_M_P", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                Me.TextBox2.Text = MiDataTable.Rows(0).Item(1).ToString
                Me.TextBox8.Text = MiDataTable.Rows(0).Item(2).ToString
                Me.TextBox30.Text = MiDataTable.Rows(0).Item(3).ToString
                Me.TextBox3.Text = Mid(MiDataTable.Rows(0).Item(4).ToString, 1, 10)
                Me.TextBox9.Text = MiDataTable.Rows(0).Item(5).ToString
                Me.TextBox7.Text = MiDataTable.Rows(0).Item(6).ToString
                Me.TextBox31.Text = MiDataTable.Rows(0).Item(7).ToString
                Me.ComboBox9.Text = MiDataTable.Rows(0).Item(8).ToString
                Me.ComboBox21.Text = MiDataTable.Rows(0).Item(9).ToString
                Me.ComboBox1.Text = MiDataTable.Rows(0).Item(10).ToString
            Else
                Me.TextBox2.Text = ""
                Me.TextBox8.Text = ""
                Me.TextBox30.Text = ""
                Me.TextBox3.Text = ""
                Me.TextBox9.Text = ""
                Me.TextBox7.Text = ""
                Me.TextBox31.Text = ""
                Me.ComboBox9.Text = "."
                Me.ComboBox21.Text = "."
                Me.ComboBox1.Text = "."
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim DE, SE As String
    Private Sub BUSCAR_UBICACION_Y_CARGO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT ID_M_P, D_N1, D_N2, D_N3, D_N4, D_N5, D_N6, D_N7, D_N8, D_CARGO_ES, ID_SITUACION FROM [dbo].[VISTA MAESTRO DE CARGOS SIT] WHERE ID_M_P = '" & h1ID_M_P & "' AND ID_SITUACION = 1 ORDER BY ID_M_P", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                'Me.TextBox11.Text = MiDataTable.Rows(0).Item(1).ToString
                'Me.TextBox12.Text = MiDataTable.Rows(0).Item(9).ToString



                '---------------------------------------------------------------------------------- SERVICIO
                If MiDataTable.Rows(0).Item(8).ToString <> "" Then  'NIVEL 8
                    SE = MiDataTable.Rows(0).Item(8).ToString
                    GoTo SERVICIO
                Else
                    If MiDataTable.Rows(0).Item(7).ToString <> "" Then  'NIVEL 7
                        SE = MiDataTable.Rows(0).Item(7).ToString
                        GoTo SERVICIO
                    Else
                        If MiDataTable.Rows(0).Item(6).ToString <> "" Then   'NIVEL 6
                            SE = MiDataTable.Rows(0).Item(6).ToString
                            GoTo SERVICIO
                        Else
                            If MiDataTable.Rows(0).Item(5).ToString <> "" Then  'NIVEL 5
                                SE = MiDataTable.Rows(0).Item(5).ToString
                                GoTo SERVICIO
                            Else
                                If MiDataTable.Rows(0).Item(4).ToString <> "" Then  'NIVEL 4
                                    SE = MiDataTable.Rows(0).Item(4).ToString
                                    GoTo SERVICIO
                                Else
                                    If MiDataTable.Rows(0).Item(3).ToString <> "" Then  'NIVEL 3
                                        SE = MiDataTable.Rows(0).Item(3).ToString
                                        GoTo SERVICIO
                                    Else
                                        If MiDataTable.Rows(0).Item(2).ToString <> "" Then  'NIVEL 2
                                            SE = MiDataTable.Rows(0).Item(2).ToString
                                            GoTo SERVICIO
                                        Else
                                            If MiDataTable.Rows(0).Item(1).ToString <> "" Then  'NIVEL 1
                                                SE = MiDataTable.Rows(0).Item(1).ToString
                                                GoTo SERVICIO
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
SERVICIO:
                '---------------------------------------------------------------------------------- DEPARTAMENTO
                If MiDataTable.Rows(0).Item(8).ToString <> "" Then  'NIVEL 8
                    DE = MiDataTable.Rows(0).Item(8).ToString
                    If DE <> SE Then
                        GoTo DEPARTAMENTO
                    Else
                        GoTo LEE_1
                    End If
                Else
LEE_1:
                    If MiDataTable.Rows(0).Item(7).ToString <> "" Then  'NIVEL 7
                        DE = MiDataTable.Rows(0).Item(7).ToString
                        If DE <> SE Then
                            GoTo DEPARTAMENTO
                        Else
                            GoTo LEE_2
                        End If
                    Else
LEE_2:
                        If MiDataTable.Rows(0).Item(6).ToString <> "" Then  'NIVEL 6
                            DE = MiDataTable.Rows(0).Item(6).ToString
                            If DE <> SE Then
                                GoTo DEPARTAMENTO
                            Else
                                GoTo LEE_3
                            End If
                        Else
LEE_3:
                            If MiDataTable.Rows(0).Item(5).ToString <> "" Then  'NIVEL 5
                                DE = MiDataTable.Rows(0).Item(5).ToString
                                If DE <> SE Then
                                    GoTo DEPARTAMENTO
                                Else
                                    GoTo LEE_4
                                End If
                            Else
LEE_4:
                                If MiDataTable.Rows(0).Item(4).ToString <> "" Then  'NIVEL 4
                                    DE = MiDataTable.Rows(0).Item(4).ToString
                                    If DE <> SE Then
                                        GoTo DEPARTAMENTO
                                    Else
                                        GoTo LEE_5
                                    End If
                                Else
LEE_5:
                                    If MiDataTable.Rows(0).Item(3).ToString <> "" Then  'NIVEL 3
                                        DE = MiDataTable.Rows(0).Item(3).ToString
                                        If DE <> SE Then
                                            GoTo DEPARTAMENTO
                                        Else
                                            GoTo LEE_6
                                        End If
                                    Else
LEE_6:
                                        If MiDataTable.Rows(0).Item(2).ToString <> "" Then  'NIVEL 2
                                            DE = MiDataTable.Rows(0).Item(2).ToString
                                            If DE <> SE Then
                                                GoTo DEPARTAMENTO
                                            Else
                                                GoTo LEE_7
                                            End If
                                        Else
LEE_7:
                                            If MiDataTable.Rows(0).Item(1).ToString <> "" Then  'NIVEL 1
                                                DE = MiDataTable.Rows(0).Item(1).ToString
                                                If DE <> SE Then
                                                    GoTo DEPARTAMENTO
                                                Else
                                                    GoTo LEE_8
                                                End If
                                            End If
LEE_8:
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
DEPARTAMENTO:
                Me.TextBox11.Text = DE
                Me.TextBox12.Text = SE

                Me.TextBox13.Text = MiDataTable.Rows(0).Item(9).ToString
                DATO_EN_RRHH = True
            Else
                Me.TextBox11.Text = ""
                Me.TextBox12.Text = ""
                Me.TextBox13.Text = ""
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_4()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE HIGIENE DISPOSICION FINAL Y ANEXOS]")
        DGV4.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE HIGIENE DISPOSICION FINAL Y ANEXOS]")
        Me.Label1.Text = DGV4.RowCount & " Registro(s)"
        If DGV4.RowCount <> 0 Then
            Me.Button8.Enabled = True
            Me.Button5.Enabled = True
            Me.Button4.Enabled = True
            Me.DGV4.Enabled = True
        Else
            Me.Button8.Enabled = True
            Me.Button5.Enabled = False
            Me.Button4.Enabled = False
            Me.DGV4.Enabled = False
        End If
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub SELECCIONAR_DATAGRIDVIEW_4()
        Dim I As Integer
        For I = 0 To Me.DGV4.RowCount - 1
            If Me.DGV4.Item(0, I).Value = h4ID_DFA Then
                Me.DGV4.Rows(I).Selected = True
                Me.DGV4.CurrentCell = Me.DGV4.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_4()
        Me.DGV4.Columns(0).HeaderText = "Id"
        Me.DGV4.Columns(0).Width = 30
        Me.DGV4.Columns(0).Visible = True
        Me.DGV4.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV4.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV4.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Me.DGV4.Columns(1).HeaderText = "ID_M_IT"
        Me.DGV4.Columns(1).Width = 0
        Me.DGV4.Columns(1).Visible = False

        Me.DGV4.Columns(2).HeaderText = "ID_M_P"
        Me.DGV4.Columns(2).Width = 0
        Me.DGV4.Columns(2).Visible = False

        Me.DGV4.Columns(3).HeaderText = "ID_AMC"
        Me.DGV4.Columns(3).Width = 0
        Me.DGV4.Columns(3).Visible = False

        Me.DGV4.Columns(4).HeaderText = "Acciones Mejoras Continuas"
        Me.DGV4.Columns(4).Width = 180
        Me.DGV4.Columns(4).Visible = True
        Me.DGV4.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV4.Columns(5).HeaderText = "Conclusiones"
        Me.DGV4.Columns(5).Width = 140
        Me.DGV4.Columns(5).Visible = True
        Me.DGV4.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV4.Columns(6).HeaderText = "Ley y Normativas"
        Me.DGV4.Columns(6).Width = 120
        Me.DGV4.Columns(6).Visible = True
        Me.DGV4.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV4.Columns(7).HeaderText = "Registro Admi. Emergencia"
        Me.DGV4.Columns(7).Width = 120
        Me.DGV4.Columns(7).Visible = True
        Me.DGV4.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV4.Columns(8).HeaderText = "Registro Resumen Clinico"
        Me.DGV4.Columns(8).Width = 120
        Me.DGV4.Columns(8).Visible = True
        Me.DGV4.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV4.Columns(9).HeaderText = "RC" : Me.DGV4.Columns(9).Width = 0 : Me.DGV4.Columns(9).Visible = False
        Me.DGV4.Columns(10).HeaderText = "AMC" : Me.DGV4.Columns(10).Width = 0 : Me.DGV4.Columns(10).Visible = False
        Me.DGV4.Columns(11).HeaderText = "AMCMA" : Me.DGV4.Columns(11).Width = 0 : Me.DGV4.Columns(11).Visible = False
        Me.DGV4.Columns(12).HeaderText = "RT" : Me.DGV4.Columns(12).Width = 0 : Me.DGV4.Columns(12).Visible = False
        Me.DGV4.Columns(13).HeaderText = "OBS" : Me.DGV4.Columns(13).Width = 0 : Me.DGV4.Columns(13).Visible = False
        Me.DGV4.Columns(14).HeaderText = "ANEXOS" : Me.DGV4.Columns(14).Width = 0 : Me.DGV4.Columns(14).Visible = False

    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_3()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE HIGIENE CAUSAS FACTORES Y GRAVEDAD]")
        DGV3.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE HIGIENE CAUSAS FACTORES Y GRAVEDAD]")
        Me.Label81.Text = DGV3.RowCount & " Registro(s)"
        If DGV3.RowCount <> 0 Then
            Me.Button17.Enabled = True
            Me.Button16.Enabled = True
            Me.Button15.Enabled = True
            Me.DGV3.Enabled = True
        Else
            Me.Button17.Enabled = True
            Me.Button16.Enabled = False
            Me.Button15.Enabled = False
            Me.DGV3.Enabled = False
        End If
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub SELECCIONAR_DATAGRIDVIEW_3()
        Dim I As Integer
        For I = 0 To Me.DGV3.RowCount - 1
            If Me.DGV3.Item(0, I).Value = h3ID_CFG Then
                Me.DGV3.Rows(I).Selected = True
                Me.DGV3.CurrentCell = Me.DGV3.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_3()
        Me.DGV3.Columns(0).HeaderText = "Id"
        Me.DGV3.Columns(0).Width = 30
        Me.DGV3.Columns(0).Visible = True
        Me.DGV3.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV3.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV3.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Me.DGV3.Columns(1).HeaderText = "ID_M_IT"
        Me.DGV3.Columns(1).Width = 0
        Me.DGV3.Columns(1).Visible = False

        Me.DGV3.Columns(2).HeaderText = "ID_M_P"
        Me.DGV3.Columns(2).Width = 0
        Me.DGV3.Columns(2).Visible = False

        Me.DGV3.Columns(3).HeaderText = "ID_OSP"
        Me.DGV3.Columns(3).Width = 0
        Me.DGV3.Columns(3).Visible = False

        Me.DGV3.Columns(4).HeaderText = "Organizativas/ Sist. Prevencion"
        Me.DGV3.Columns(4).Width = 180
        Me.DGV3.Columns(4).Visible = True
        Me.DGV3.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(5).HeaderText = "ID_FP"
        Me.DGV3.Columns(5).Width = 0
        Me.DGV3.Columns(5).Visible = False

        Me.DGV3.Columns(6).HeaderText = "Factores Personales"
        Me.DGV3.Columns(6).Width = 100
        Me.DGV3.Columns(6).Visible = True
        Me.DGV3.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(7).HeaderText = "ID_FT"
        Me.DGV3.Columns(7).Width = 0
        Me.DGV3.Columns(7).Visible = False

        Me.DGV3.Columns(8).HeaderText = "Factores Trabajo"
        Me.DGV3.Columns(8).Width = 100
        Me.DGV3.Columns(8).Visible = True
        Me.DGV3.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(9).HeaderText = "ID_AI"
        Me.DGV3.Columns(9).Width = 0
        Me.DGV3.Columns(9).Visible = False

        Me.DGV3.Columns(10).HeaderText = "Actos Inseguros"
        Me.DGV3.Columns(10).Width = 100
        Me.DGV3.Columns(10).Visible = True
        Me.DGV3.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(11).HeaderText = "ID_CI"
        Me.DGV3.Columns(11).Width = 0
        Me.DGV3.Columns(11).Visible = False

        Me.DGV3.Columns(12).HeaderText = "Condiciones Inseguras"
        Me.DGV3.Columns(12).Width = 100
        Me.DGV3.Columns(12).Visible = True
        Me.DGV3.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(13).HeaderText = "ID_FR"
        Me.DGV3.Columns(13).Width = 0
        Me.DGV3.Columns(13).Visible = False

        Me.DGV3.Columns(14).HeaderText = "Factores Riesgos"
        Me.DGV3.Columns(14).Width = 100
        Me.DGV3.Columns(14).Visible = True
        Me.DGV3.Columns(14).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(15).HeaderText = "ID_NG"
        Me.DGV3.Columns(15).Width = 0
        Me.DGV3.Columns(15).Visible = False

        Me.DGV3.Columns(16).HeaderText = "Nivel Gravedad"
        Me.DGV3.Columns(16).Width = 0
        Me.DGV3.Columns(16).Visible = False
        Me.DGV3.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(17).HeaderText = "FG" : Me.DGV3.Columns(17).Width = 0 : Me.DGV3.Columns(17).Visible = False
        Me.DGV3.Columns(18).HeaderText = "FSE" : Me.DGV3.Columns(18).Width = 0 : Me.DGV3.Columns(18).Visible = False
        Me.DGV3.Columns(19).HeaderText = "RACC" : Me.DGV3.Columns(19).Width = 0 : Me.DGV3.Columns(19).Visible = False
        Me.DGV3.Columns(20).HeaderText = "NLG" : Me.DGV3.Columns(20).Width = 0 : Me.DGV3.Columns(20).Visible = False
        Me.DGV3.Columns(21).HeaderText = "AC" : Me.DGV3.Columns(21).Width = 0 : Me.DGV3.Columns(21).Visible = False
        Me.DGV3.Columns(22).HeaderText = "NG" : Me.DGV3.Columns(22).Width = 0 : Me.DGV3.Columns(22).Visible = False
        Me.DGV3.Columns(23).HeaderText = "SUBS" : Me.DGV3.Columns(23).Width = 0 : Me.DGV3.Columns(23).Visible = False
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_2()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE HIGIENE DATOS DEL EVENTO]")
        DGV2.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE HIGIENE DATOS DEL EVENTO]")
        Me.Label80.Text = DGV2.RowCount & " Registro(s)"
        If DGV1.RowCount <> 0 Then
            Me.Button14.Enabled = True
            Me.Button13.Enabled = True
            Me.Button12.Enabled = True
            Me.DGV2.Enabled = True
        Else
            Me.Button14.Enabled = True
            Me.Button13.Enabled = False
            Me.Button12.Enabled = False
            Me.DGV2.Enabled = False
        End If
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub SELECCIONAR_DATAGRIDVIEW_2()
        Dim I As Integer
        For I = 0 To Me.DGV2.RowCount - 1
            If Me.DGV2.Item(0, I).Value = mieID_M_EVENTO Then
                Me.DGV2.Rows(I).Selected = True
                Me.DGV2.CurrentCell = Me.DGV2.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_2()
        Me.DGV2.Columns(0).HeaderText = "Id"
        Me.DGV2.Columns(0).Width = 30
        Me.DGV2.Columns(0).Visible = True
        Me.DGV2.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV2.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV2.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Me.DGV2.Columns(1).HeaderText = "ID_M_IT"
        Me.DGV2.Columns(1).Width = 0
        Me.DGV2.Columns(1).Visible = False

        Me.DGV2.Columns(2).HeaderText = "ID_M_P"
        Me.DGV2.Columns(2).Width = 0
        Me.DGV2.Columns(2).Visible = False

        Me.DGV2.Columns(3).HeaderText = "Descripcion Accidente"
        Me.DGV2.Columns(3).Width = 200
        Me.DGV2.Columns(3).Visible = True
        Me.DGV2.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV2.Columns(4).HeaderText = "ID_T_ACCIDENTE"
        Me.DGV2.Columns(4).Width = 0
        Me.DGV2.Columns(4).Visible = False

        Me.DGV2.Columns(5).HeaderText = "Tipo Accidente"
        Me.DGV2.Columns(5).Width = 80
        Me.DGV2.Columns(5).Visible = True
        Me.DGV2.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV2.Columns(6).HeaderText = "Lugar Accidente"
        Me.DGV2.Columns(6).Width = 80
        Me.DGV2.Columns(6).Visible = True
        Me.DGV2.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV2.Columns(7).HeaderText = "Fecha Accidente"
        Me.DGV2.Columns(7).Width = 80
        Me.DGV2.Columns(7).Visible = True
        Me.DGV2.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV2.Columns(8).HeaderText = "Hora Accidente"
        Me.DGV2.Columns(8).Width = 80
        Me.DGV2.Columns(8).Visible = True
        Me.DGV2.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV2.Columns(8).DefaultCellStyle.Format = "t"

        Me.DGV2.Columns(9).HeaderText = "Fecha Reporte"
        Me.DGV2.Columns(9).Width = 80
        Me.DGV2.Columns(9).Visible = True
        Me.DGV2.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV2.Columns(10).HeaderText = "Hora Reporte"
        Me.DGV2.Columns(10).Width = 80
        Me.DGV2.Columns(10).Visible = True
        Me.DGV2.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV2.Columns(10).DefaultCellStyle.Format = "t"

        Me.DGV2.Columns(11).HeaderText = "Horas Laboradas"
        Me.DGV2.Columns(11).Width = 0
        Me.DGV2.Columns(11).Visible = False
        Me.DGV2.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV2.Columns(12).HeaderText = "Horas Perdidas"
        Me.DGV2.Columns(12).Width = 0
        Me.DGV2.Columns(12).Visible = False
        Me.DGV2.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV2.Columns(13).HeaderText = "ID_AMG"
        Me.DGV2.Columns(13).Width = 0
        Me.DGV2.Columns(13).Visible = False

        Me.DGV2.Columns(14).HeaderText = "Agente Material General"
        Me.DGV2.Columns(14).Width = 0
        Me.DGV2.Columns(14).Visible = False
        Me.DGV2.Columns(14).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV2.Columns(15).HeaderText = "ID_AMSE"
        Me.DGV2.Columns(15).Width = 0
        Me.DGV2.Columns(15).Visible = False

        Me.DGV2.Columns(16).HeaderText = "Agente Material Sub Estandar"
        Me.DGV2.Columns(16).Width = 0
        Me.DGV2.Columns(16).Visible = False
        Me.DGV2.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV2.Columns(17).HeaderText = "ID_FO"
        Me.DGV2.Columns(17).Width = 0
        Me.DGV2.Columns(17).Visible = False

        Me.DGV2.Columns(18).HeaderText = "Forma Ocurrencia"
        Me.DGV2.Columns(18).Width = 0
        Me.DGV2.Columns(18).Visible = False
        Me.DGV2.Columns(18).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV2.Columns(19).HeaderText = "ID_PCL"
        Me.DGV2.Columns(19).Width = 0
        Me.DGV2.Columns(19).Visible = False

        Me.DGV2.Columns(20).HeaderText = "Parte Cuerpo Lesionada"
        Me.DGV2.Columns(20).Width = 0
        Me.DGV2.Columns(20).Visible = False
        Me.DGV2.Columns(20).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV2.Columns(21).HeaderText = "ID_NL"
        Me.DGV2.Columns(21).Width = 0
        Me.DGV2.Columns(21).Visible = False

        Me.DGV2.Columns(22).HeaderText = "Naturaleza Lesion"
        Me.DGV2.Columns(22).Width = 0
        Me.DGV2.Columns(22).Visible = False
        Me.DGV2.Columns(22).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV2.Columns(23).HeaderText = "Dias Subsidio"
        Me.DGV2.Columns(23).Width = 0
        Me.DGV2.Columns(23).Visible = False
        Me.DGV2.Columns(23).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV2.Columns(24).HeaderText = "Delcaracion Testigo 1"
        Me.DGV2.Columns(24).Width = 0
        Me.DGV2.Columns(24).Visible = False
        Me.DGV2.Columns(24).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV2.Columns(25).HeaderText = "Delcaracion Testigo 2"
        Me.DGV2.Columns(25).Width = 0
        Me.DGV2.Columns(25).Visible = False
        Me.DGV2.Columns(25).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV2.Columns(26).HeaderText = "" : Me.DGV2.Columns(26).Width = 0 : Me.DGV2.Columns(26).Visible = False
        Me.DGV2.Columns(27).HeaderText = "" : Me.DGV2.Columns(27).Width = 0 : Me.DGV2.Columns(27).Visible = False
        Me.DGV2.Columns(28).HeaderText = "" : Me.DGV2.Columns(28).Width = 0 : Me.DGV2.Columns(28).Visible = False
        Me.DGV2.Columns(29).HeaderText = "" : Me.DGV2.Columns(29).Width = 0 : Me.DGV2.Columns(29).Visible = False
        Me.DGV2.Columns(30).HeaderText = "" : Me.DGV2.Columns(30).Width = 0 : Me.DGV2.Columns(30).Visible = False
        Me.DGV2.Columns(31).HeaderText = "" : Me.DGV2.Columns(31).Width = 0 : Me.DGV2.Columns(31).Visible = False
        Me.DGV2.Columns(32).HeaderText = "" : Me.DGV2.Columns(32).Width = 0 : Me.DGV2.Columns(32).Visible = False
        Me.DGV2.Columns(33).HeaderText = "" : Me.DGV2.Columns(33).Width = 0 : Me.DGV2.Columns(33).Visible = False
        Me.DGV2.Columns(34).HeaderText = "" : Me.DGV2.Columns(34).Width = 0 : Me.DGV2.Columns(34).Visible = False
        Me.DGV2.Columns(35).HeaderText = "" : Me.DGV2.Columns(35).Width = 0 : Me.DGV2.Columns(35).Visible = False
        Me.DGV2.Columns(36).HeaderText = "" : Me.DGV2.Columns(36).Width = 0 : Me.DGV2.Columns(36).Visible = False
        Me.DGV2.Columns(37).HeaderText = "" : Me.DGV2.Columns(37).Width = 0 : Me.DGV2.Columns(37).Visible = False
        Me.DGV2.Columns(38).HeaderText = "" : Me.DGV2.Columns(38).Width = 0 : Me.DGV2.Columns(38).Visible = False
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_1()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE HIGIENE INFORMACION DEL ACCIDENTE]")
        DGV1.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE HIGIENE INFORMACION DEL ACCIDENTE]")
        Me.Label61.Text = DGV1.RowCount & " Registro(s)"
        If DGV1.RowCount <> 0 Then
            Me.Button11.Enabled = True
            Me.Button10.Enabled = True
            Me.Button9.Enabled = True
            Me.DGV1.Enabled = True
        Else
            Me.Button11.Enabled = True
            Me.Button10.Enabled = False
            Me.Button9.Enabled = False
            Me.DGV1.Enabled = False
        End If
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub SELECCIONAR_DATAGRIDVIEW_1()
        Dim I As Integer
        For I = 0 To Me.DGV1.RowCount - 1
            If Me.DGV1.Item(0, I).Value = h2ID_M_IT Then
                Me.DGV1.Rows(I).Selected = True
                Me.DGV1.CurrentCell = Me.DGV1.Item(0, I)
                Exit Sub
            End If
        Next
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

        Me.DGV1.Columns(2).HeaderText = "Parte Accidente"
        Me.DGV1.Columns(2).Width = 70
        Me.DGV1.Columns(2).Visible = True
        Me.DGV1.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(3).HeaderText = "Reporte Inss"
        Me.DGV1.Columns(3).Width = 70
        Me.DGV1.Columns(3).Visible = True
        Me.DGV1.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(4).HeaderText = "Reporte Mitrab"
        Me.DGV1.Columns(4).Width = 70
        Me.DGV1.Columns(4).Visible = True
        Me.DGV1.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(5).HeaderText = "Edad"
        Me.DGV1.Columns(5).Width = 0
        Me.DGV1.Columns(5).Visible = False
        Me.DGV1.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(6).HeaderText = "Accidente Anterior"
        Me.DGV1.Columns(6).Width = 0
        Me.DGV1.Columns(6).Visible = False
        Me.DGV1.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(7).HeaderText = "Departamento"
        Me.DGV1.Columns(7).Width = 0
        Me.DGV1.Columns(7).Visible = False
        Me.DGV1.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(8).HeaderText = "Servicio"
        Me.DGV1.Columns(8).Width = 0
        Me.DGV1.Columns(8).Visible = False
        Me.DGV1.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(9).HeaderText = "Cargo"
        Me.DGV1.Columns(9).Width = 0
        Me.DGV1.Columns(9).Visible = False
        Me.DGV1.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(10).HeaderText = "Antigue. Empresa"
        Me.DGV1.Columns(10).Width = 0
        Me.DGV1.Columns(10).Visible = False
        Me.DGV1.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(11).HeaderText = "Antigue. Puesto"
        Me.DGV1.Columns(11).Width = 0
        Me.DGV1.Columns(11).Visible = False
        Me.DGV1.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

    End Sub
    Private Sub DGV1_SelectionChanged(sender As Object, e As EventArgs) Handles DGV1.SelectionChanged
        If Me.DGV1.RowCount > 0 Then
            For I = 0 To Me.DGV1.RowCount - 1
                If DGV1.Rows(I).Selected = True Then
                    h2ID_M_IT = Me.DGV1.Item(0, I).Value
                    h2ID_M_P = Me.DGV1.Item(1, I).Value
                    h2NO_PARTE_ACC = Me.DGV1.Item(2, I).Value
                    h2NO_REP_NSS = Me.DGV1.Item(3, I).Value
                    h2NO_REP_MITRAB = Me.DGV1.Item(4, I).Value
                    h2EDAD = Me.DGV1.Item(5, I).Value
                    h2ACCIDENTE_ANTERIOR = Me.DGV1.Item(6, I).Value
                    h2DEPARTAMENTO = Me.DGV1.Item(7, I).Value
                    h2SERVICIO = Me.DGV1.Item(8, I).Value
                    h2CARGO = Me.DGV1.Item(9, I).Value
                    h2ANT_EMPRESA_MESES = Me.DGV1.Item(10, I).Value
                    h2ANT_PUESTO_MESES = Me.DGV1.Item(11, I).Value
                    '-------------------------------------------------------
                    CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HIGIENE DATOS DEL EVENTO] WHERE ID_M_P = " & h1ID_M_P & " AND ID_M_IT = " & h2ID_M_IT & " ORDER BY FECHA_ACC"
                    Call CARGAR_DATAGRIDVIEW_2()
                    Call ARMAR_DATAGRIDVIEW_2()
                    '-------------------------------------------------------
                    CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HIGIENE CAUSAS FACTORES Y GRAVEDAD] WHERE ID_M_P = " & h1ID_M_P & " AND ID_M_IT = " & h2ID_M_IT & " ORDER BY ID_CFG"
                    Call CARGAR_DATAGRIDVIEW_3()
                    Call ARMAR_DATAGRIDVIEW_3()
                    '-------------------------------------------------------
                    CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HIGIENE DISPOSICION FINAL Y ANEXOS] WHERE ID_M_P = " & h1ID_M_P & " AND ID_M_IT = " & h2ID_M_IT & " ORDER BY ID_DFA"
                    Call CARGAR_DATAGRIDVIEW_4()
                    Call ARMAR_DATAGRIDVIEW_4()
                    '-------------------------------------------------------
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub DGV1_Click(sender As Object, e As EventArgs) Handles DGV1.Click
        If Me.DGV1.RowCount > 0 Then
            For I = 0 To Me.DGV1.RowCount - 1
                If DGV1.Rows(I).Selected = True Then
                    h2ID_M_IT = Me.DGV1.Item(0, I).Value
                    h2ID_M_P = Me.DGV1.Item(1, I).Value
                    h2NO_PARTE_ACC = Me.DGV1.Item(2, I).Value
                    h2NO_REP_NSS = Me.DGV1.Item(3, I).Value
                    h2NO_REP_MITRAB = Me.DGV1.Item(4, I).Value
                    h2EDAD = Me.DGV1.Item(5, I).Value
                    h2ACCIDENTE_ANTERIOR = Me.DGV1.Item(6, I).Value
                    h2DEPARTAMENTO = Me.DGV1.Item(7, I).Value
                    h2SERVICIO = Me.DGV1.Item(8, I).Value
                    h2CARGO = Me.DGV1.Item(9, I).Value
                    h2ANT_EMPRESA_MESES = Me.DGV1.Item(10, I).Value
                    h2ANT_PUESTO_MESES = Me.DGV1.Item(11, I).Value
                    '-------------------------------------------------------
                    CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HIGIENE DATOS DEL EVENTO] WHERE ID_M_P = " & h1ID_M_P & " AND ID_M_IT = " & h2ID_M_IT & " ORDER BY FECHA_ACC"
                    Call CARGAR_DATAGRIDVIEW_2()
                    Call ARMAR_DATAGRIDVIEW_2()
                    '-------------------------------------------------------
                    CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HIGIENE CAUSAS FACTORES Y GRAVEDAD] WHERE ID_M_P = " & h1ID_M_P & " AND ID_M_IT = " & h2ID_M_IT & " ORDER BY ID_CFG"
                    Call CARGAR_DATAGRIDVIEW_3()
                    Call ARMAR_DATAGRIDVIEW_3()
                    '-------------------------------------------------------
                    CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HIGIENE DISPOSICION FINAL Y ANEXOS] WHERE ID_M_P = " & h1ID_M_P & " AND ID_M_IT = " & h2ID_M_IT & " ORDER BY ID_DFA"
                    Call CARGAR_DATAGRIDVIEW_4()
                    Call ARMAR_DATAGRIDVIEW_4()
                    '-------------------------------------------------------
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Call VERIFICAR_ACCESOS("195") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Accidentes_Laborales_Add_Inf_Acc.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HIGIENE INFORMACION DEL ACCIDENTE] WHERE ID_M_P = " & h1ID_M_P & " ORDER BY ID_M_P, NO_PARTE_ACC"
            Call CARGAR_DATAGRIDVIEW_1()
            Call ARMAR_DATAGRIDVIEW_1()
            Call SELECCIONAR_DATAGRIDVIEW_1()
        End If
    End Sub
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Call VERIFICAR_ACCESOS("196") : If HAY_ACCESO = False Then : Exit Sub : End If
        If h2ID_M_IT = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Accidentes_Laborales_Upd_Inf_Acc.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HIGIENE INFORMACION DEL ACCIDENTE] WHERE ID_M_P = " & h1ID_M_P & " ORDER BY ID_M_P, NO_PARTE_ACC"
            Call CARGAR_DATAGRIDVIEW_1()
            Call ARMAR_DATAGRIDVIEW_1()
            Call SELECCIONAR_DATAGRIDVIEW_1()
        End If
    End Sub
    Private Sub DGV1_DoubleClick(sender As Object, e As EventArgs) Handles DGV1.DoubleClick
        Call VERIFICAR_ACCESOS("196") : If HAY_ACCESO = False Then : Exit Sub : End If
        If h2ID_M_IT = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Accidentes_Laborales_Upd_Inf_Acc.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HIGIENE INFORMACION DEL ACCIDENTE] WHERE ID_M_P = " & h1ID_M_P & " ORDER BY ID_M_P, NO_PARTE_ACC"
            Call CARGAR_DATAGRIDVIEW_1()
            Call ARMAR_DATAGRIDVIEW_1()
            Call SELECCIONAR_DATAGRIDVIEW_1()
        End If
    End Sub
    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Call VERIFICAR_ACCESOS("198") : If HAY_ACCESO = False Then : Exit Sub : End If
        If h2ID_M_IT = 0 Then
            MsgBox("Debe seleccionar el registro de la Información del Accidente del que desea Visualizar los Eventos", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Datos_Eventos.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HIGIENE DATOS DEL EVENTO] WHERE ID_M_P = " & h1ID_M_P & " AND ID_M_IT = " & h2ID_M_IT & " ORDER BY FECHA_ACC"
            Call CARGAR_DATAGRIDVIEW_2()
            Call ARMAR_DATAGRIDVIEW_2()
            Call SELECCIONAR_DATAGRIDVIEW_2()
        End If
    End Sub
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Call VERIFICAR_ACCESOS("199") : If HAY_ACCESO = False Then : Exit Sub : End If
        If h2ID_M_IT = 0 Then
            MsgBox("Debe seleccionar el registro de la Información del Accidente del que desea Visualizar los Eventos", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        If mieID_M_EVENTO = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Accidentes_Laborales_Upd_Dat_Evento.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HIGIENE DATOS DEL EVENTO] WHERE ID_M_P = " & h1ID_M_P & " AND ID_M_IT = " & h2ID_M_IT & " ORDER BY FECHA_ACC"
            Call CARGAR_DATAGRIDVIEW_2()
            Call ARMAR_DATAGRIDVIEW_2()
            Call SELECCIONAR_DATAGRIDVIEW_2()
        End If
    End Sub
    Private Sub DGV2_Click(sender As Object, e As EventArgs) Handles DGV2.Click
        If DGV2.SelectedCells.Count <> 0 Then
            Dim FILA = DGV2.CurrentRow.Index
            If Not IsDBNull(Me.DGV2.Item(0, FILA).Value) Then : mieID_M_EVENTO = Me.DGV2.Item(0, FILA).Value : Else : mieID_M_EVENTO = 0 : End If
            If Not IsDBNull(Me.DGV2.Item(1, FILA).Value) Then : mieID_M_IT = Me.DGV2.Item(1, FILA).Value : Else : mieID_M_IT = 0 : End If
            If Not IsDBNull(Me.DGV2.Item(1, FILA).Value) Then : mieID_M_P = Me.DGV2.Item(2, FILA).Value : Else : mieID_M_P = 0 : End If
            If Not IsDBNull(Me.DGV2.Item(3, FILA).Value) Then : mieDESCRIPCION_ACC = Me.DGV2.Item(3, FILA).Value : Else : mieDESCRIPCION_ACC = "" : End If
            If Not IsDBNull(Me.DGV2.Item(4, FILA).Value) Then : mieID_T_ACCIDENTE = Me.DGV2.Item(4, FILA).Value : Else : mieID_T_ACCIDENTE = 0 : End If
            If Not IsDBNull(Me.DGV2.Item(5, FILA).Value) Then : mieD_T_ACCIDENTE = Me.DGV2.Item(5, FILA).Value : Else : mieD_T_ACCIDENTE = "" : End If
            If Not IsDBNull(Me.DGV2.Item(6, FILA).Value) Then : mieLUGAR_ACC = Me.DGV2.Item(6, FILA).Value : Else : mieLUGAR_ACC = "" : End If
            If Not IsDBNull(Me.DGV2.Item(7, FILA).Value) Then : mieFECHA_ACC = Me.DGV2.Item(7, FILA).Value : Else : mieFECHA_ACC = "" : End If
            If Not IsDBNull(Me.DGV2.Item(8, FILA).Value) Then : mieHORA_ACC = Me.DGV2.Item(8, FILA).Value : Else : mieHORA_ACC = "" : End If
            If Not IsDBNull(Me.DGV2.Item(9, FILA).Value) Then : mieFECHA_REPORTE = Me.DGV2.Item(9, FILA).Value : Else : mieFECHA_REPORTE = "" : End If
            If Not IsDBNull(Me.DGV2.Item(10, FILA).Value) Then : mieHORA_REPORTE = Me.DGV2.Item(10, FILA).Value : Else : mieHORA_REPORTE = "" : End If

            If Not IsDBNull(Me.DGV2.Item(11, FILA).Value) Then : mieHORAS_LABORADAS = Me.DGV2.Item(11, FILA).Value : Else : mieHORAS_LABORADAS = "" : End If
            If Not IsDBNull(Me.DGV2.Item(12, FILA).Value) Then : mieHORAS_PERDIDAS = Me.DGV2.Item(12, FILA).Value : Else : mieHORAS_PERDIDAS = "" : End If
            If Not IsDBNull(Me.DGV2.Item(13, FILA).Value) Then : mieID_AMG = Me.DGV2.Item(13, FILA).Value : Else : mieID_AMG = 0 : End If
            If Not IsDBNull(Me.DGV2.Item(14, FILA).Value) Then : mieD_AMG = Me.DGV2.Item(14, FILA).Value : Else : mieD_AMG = "" : End If
            If Not IsDBNull(Me.DGV2.Item(15, FILA).Value) Then : mieID_AMSE = Me.DGV2.Item(15, FILA).Value : Else : mieID_AMSE = 0 : End If
            If Not IsDBNull(Me.DGV2.Item(16, FILA).Value) Then : mieD_AMSE = Me.DGV2.Item(16, FILA).Value : Else : mieD_AMSE = "" : End If
            If Not IsDBNull(Me.DGV2.Item(17, FILA).Value) Then : mieID_FO = Me.DGV2.Item(17, FILA).Value : Else : mieID_FO = 0 : End If
            If Not IsDBNull(Me.DGV2.Item(18, FILA).Value) Then : mieD_FO = Me.DGV2.Item(18, FILA).Value : Else : mieD_FO = "" : End If
            If Not IsDBNull(Me.DGV2.Item(19, FILA).Value) Then : mieID_PCL = Me.DGV2.Item(19, FILA).Value : Else : mieID_PCL = 0 : End If
            If Not IsDBNull(Me.DGV2.Item(20, FILA).Value) Then : mieD_PCL = Me.DGV2.Item(20, FILA).Value : Else : mieD_PCL = "" : End If

            If Not IsDBNull(Me.DGV2.Item(21, FILA).Value) Then : mieID_NL = Me.DGV2.Item(21, FILA).Value : Else : mieID_NL = 0 : End If
            If Not IsDBNull(Me.DGV2.Item(22, FILA).Value) Then : mieD_NL = Me.DGV2.Item(22, FILA).Value : Else : mieD_NL = "" : End If
            If Not IsDBNull(Me.DGV2.Item(23, FILA).Value) Then : mieDIAS_SUBSIDIO = Me.DGV2.Item(23, FILA).Value : Else : mieDIAS_SUBSIDIO = "" : End If
            If Not IsDBNull(Me.DGV2.Item(24, FILA).Value) Then : mieDECLARACION_T1 = Me.DGV2.Item(24, FILA).Value : Else : mieDECLARACION_T1 = "" : End If
            If Not IsDBNull(Me.DGV2.Item(25, FILA).Value) Then : mieDECLARACION_T2 = Me.DGV2.Item(25, FILA).Value : Else : mieDECLARACION_T2 = "" : End If
            If Not IsDBNull(Me.DGV2.Item(26, FILA).Value) Then : mieLUGAR_E_ACC = Me.DGV2.Item(26, FILA).Value : Else : mieLUGAR_E_ACC = "" : End If
            If Not IsDBNull(Me.DGV2.Item(27, FILA).Value) Then : miePCL_E = Me.DGV2.Item(27, FILA).Value : Else : miePCL_E = "" : End If
            If Not IsDBNull(Me.DGV2.Item(28, FILA).Value) Then : mieTIEMPO_O = Me.DGV2.Item(28, FILA).Value : Else : mieTIEMPO_O = "" : End If
            If Not IsDBNull(Me.DGV2.Item(29, FILA).Value) Then : mieMES = Me.DGV2.Item(29, FILA).Value : Else : mieMES = "" : End If
            If Not IsDBNull(Me.DGV2.Item(30, FILA).Value) Then : mieF_II = Me.DGV2.Item(30, FILA).Value : Else : mieF_II = "" : End If

            If Not IsDBNull(Me.DGV2.Item(31, FILA).Value) Then : mieFCI = Me.DGV2.Item(31, FILA).Value : Else : mieFCI = "" : End If
            If Not IsDBNull(Me.DGV2.Item(32, FILA).Value) Then : mieACC_ANT = Me.DGV2.Item(32, FILA).Value : Else : mieACC_ANT = "" : End If
            If Not IsDBNull(Me.DGV2.Item(33, FILA).Value) Then : mieTURNO_DE_T = Me.DGV2.Item(33, FILA).Value : Else : mieTURNO_DE_T = "" : End If
            If Not IsDBNull(Me.DGV2.Item(34, FILA).Value) Then : mieHAL = Me.DGV2.Item(34, FILA).Value : Else : mieHAL = "" : End If
            If Not IsDBNull(Me.DGV2.Item(35, FILA).Value) Then : mieSDT = Me.DGV2.Item(35, FILA).Value : Else : mieSDT = "" : End If
            If Not IsDBNull(Me.DGV2.Item(36, FILA).Value) Then : mieLQR = Me.DGV2.Item(36, FILA).Value : Else : mieLQR = "" : End If
            If Not IsDBNull(Me.DGV2.Item(37, FILA).Value) Then : mieENTREVISTADO_1 = Me.DGV2.Item(37, FILA).Value : Else : mieENTREVISTADO_1 = "" : End If
            If Not IsDBNull(Me.DGV2.Item(38, FILA).Value) Then : mieENTREVISTADO_2 = Me.DGV2.Item(38, FILA).Value : Else : mieENTREVISTADO_2 = "" : End If

        End If
    End Sub
    Private Sub DGV2_DoubleClick(sender As Object, e As EventArgs) Handles DGV2.DoubleClick
        Call VERIFICAR_ACCESOS("199") : If HAY_ACCESO = False Then : Exit Sub : End If
        If h2ID_M_IT = 0 Then
            MsgBox("Debe seleccionar el registro de la Información del Accidente del que desea Visualizar los Eventos", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        If mieID_M_EVENTO = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Accidentes_Laborales_Upd_Dat_Evento.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HIGIENE DATOS DEL EVENTO] WHERE ID_M_P = " & h1ID_M_P & " AND ID_M_IT = " & h2ID_M_IT & " ORDER BY FECHA_ACC"
            Call CARGAR_DATAGRIDVIEW_2()
            Call ARMAR_DATAGRIDVIEW_2()
            Call SELECCIONAR_DATAGRIDVIEW_2()
        End If
    End Sub
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Call VERIFICAR_ACCESOS("200") : If HAY_ACCESO = False Then : Exit Sub : End If
        If mieID_M_EVENTO = 0 Then
            MsgBox("Debe seleccionar el registro que desea Eliminar", vbInformation, "Mensaje del Sistema")
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
                MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE HIGIENE DATOS DEL EVENTO] WHERE ID_M_EVENTO = " & CInt(mieID_M_EVENTO) & ""
                MiSqlCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HIGIENE DATOS DEL EVENTO] WHERE ID_M_P = " & h1ID_M_P & " AND ID_M_IT = " & h2ID_M_IT & " ORDER BY FECHA_ACC"
            Call CARGAR_DATAGRIDVIEW_2()
            Call ARMAR_DATAGRIDVIEW_2()
        Else
            mieID_M_EVENTO = 0
        End If
    End Sub
    Private Sub DGV3_Click(sender As Object, e As EventArgs) Handles DGV3.Click
        If Me.DGV3.RowCount > 0 Then
            For I = 0 To Me.DGV3.RowCount - 1
                If DGV3.Rows(I).Selected = True Then
                    If DGV3.SelectedCells.Count <> 0 Then
                        Me.Label2.Text = DGV3.RowCount & " Registro(s)"
                        Dim FILA = DGV3.CurrentRow.Index
                        If Not IsDBNull(Me.DGV3.Item(0, FILA).Value) Then : h3ID_CFG = Me.DGV3.Item(0, FILA).Value : Else : h3ID_CFG = 0 : End If
                        If Not IsDBNull(Me.DGV3.Item(1, FILA).Value) Then : h3ID_M_IT = Me.DGV3.Item(1, FILA).Value : Else : h3ID_M_IT = 0 : End If
                        If Not IsDBNull(Me.DGV3.Item(2, FILA).Value) Then : h3ID_M_P = Me.DGV3.Item(2, FILA).Value : Else : h3ID_M_P = 0 : End If
                        If Not IsDBNull(Me.DGV3.Item(3, FILA).Value) Then : h3ID_OSP = Me.DGV3.Item(3, FILA).Value : Else : h3ID_OSP = 0 : End If
                        If Not IsDBNull(Me.DGV3.Item(4, FILA).Value) Then : h3D_OSP = Me.DGV3.Item(4, FILA).Value : Else : h3D_OSP = "" : End If
                        If Not IsDBNull(Me.DGV3.Item(5, FILA).Value) Then : h3ID_FP = Me.DGV3.Item(5, FILA).Value : Else : h3ID_FP = 0 : End If
                        If Not IsDBNull(Me.DGV3.Item(6, FILA).Value) Then : h3D_FP = Me.DGV3.Item(6, FILA).Value : Else : h3D_FP = "" : End If
                        If Not IsDBNull(Me.DGV3.Item(7, FILA).Value) Then : h3ID_FT = Me.DGV3.Item(7, FILA).Value : Else : h3ID_FT = 0 : End If
                        If Not IsDBNull(Me.DGV3.Item(8, FILA).Value) Then : h3D_FT = Me.DGV3.Item(8, FILA).Value : Else : h3D_FT = "" : End If
                        If Not IsDBNull(Me.DGV3.Item(9, FILA).Value) Then : h3ID_AI = Me.DGV3.Item(9, FILA).Value : Else : h3ID_AI = 0 : End If
                        If Not IsDBNull(Me.DGV3.Item(10, FILA).Value) Then : h3D_AI = Me.DGV3.Item(10, FILA).Value : Else : h3D_AI = "" : End If
                        If Not IsDBNull(Me.DGV3.Item(11, FILA).Value) Then : h3ID_CI = Me.DGV3.Item(11, FILA).Value : Else : h3ID_CI = 0 : End If
                        If Not IsDBNull(Me.DGV3.Item(12, FILA).Value) Then : h3D_CI = Me.DGV3.Item(12, FILA).Value : Else : h3D_CI = "" : End If
                        If Not IsDBNull(Me.DGV3.Item(13, FILA).Value) Then : h3ID_FR = Me.DGV3.Item(13, FILA).Value : Else : h3ID_FR = 0 : End If
                        If Not IsDBNull(Me.DGV3.Item(14, FILA).Value) Then : h3D_FR = Me.DGV3.Item(14, FILA).Value : Else : h3D_FR = "" : End If
                        If Not IsDBNull(Me.DGV3.Item(15, FILA).Value) Then : h3ID_NG = Me.DGV3.Item(15, FILA).Value : Else : h3ID_NG = 0 : End If
                        If Not IsDBNull(Me.DGV3.Item(16, FILA).Value) Then : h3D_NG = Me.DGV3.Item(16, FILA).Value : Else : h3D_NG = "" : End If
                        If Not IsDBNull(Me.DGV3.Item(17, FILA).Value) Then : h3FG = Me.DGV3.Item(17, FILA).Value : Else : h3FG = "" : End If
                        If Not IsDBNull(Me.DGV3.Item(18, FILA).Value) Then : h3FSE = Me.DGV3.Item(18, FILA).Value : Else : h3FSE = "" : End If
                        If Not IsDBNull(Me.DGV3.Item(19, FILA).Value) Then : h3RACC = Me.DGV3.Item(19, FILA).Value : Else : h3RACC = "" : End If
                        If Not IsDBNull(Me.DGV3.Item(20, FILA).Value) Then : h3NLG = Me.DGV3.Item(20, FILA).Value : Else : h3NLG = "" : End If
                        If Not IsDBNull(Me.DGV3.Item(21, FILA).Value) Then : h3AC = Me.DGV3.Item(21, FILA).Value : Else : h3AC = "" : End If
                        If Not IsDBNull(Me.DGV3.Item(22, FILA).Value) Then : h3NG = Me.DGV3.Item(22, FILA).Value : Else : h3NG = "" : End If
                        If Not IsDBNull(Me.DGV3.Item(23, FILA).Value) Then : h3SUBS = Me.DGV3.Item(23, FILA).Value : Else : h3SUBS = "" : End If
                    End If
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        Call VERIFICAR_ACCESOS("201") : If HAY_ACCESO = False Then : Exit Sub : End If
        If h2ID_M_IT = 0 Then
            MsgBox("Debe seleccionar el registro de la Información del Accidente del que desea Visualizar las Causas Factores y Gravedad", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Accidentes_Laborales_Add_CFG.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HIGIENE CAUSAS FACTORES Y GRAVEDAD] WHERE ID_M_P = " & h1ID_M_P & " AND ID_M_IT = " & h2ID_M_IT & " ORDER BY ID_CFG"
            Call CARGAR_DATAGRIDVIEW_3()
            Call ARMAR_DATAGRIDVIEW_3()
            Call SELECCIONAR_DATAGRIDVIEW_3()
        End If
    End Sub
    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        Call VERIFICAR_ACCESOS("202") : If HAY_ACCESO = False Then : Exit Sub : End If
        If h2ID_M_IT = 0 Then
            MsgBox("Debe seleccionar el registro de la Información del Accidente del que desea Visualizar las Causas Factores y Gravedad", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        If h3ID_CFG = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Accidentes_Laborales_Upd_CFG.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HIGIENE CAUSAS FACTORES Y GRAVEDAD] WHERE ID_M_P = " & h1ID_M_P & " AND ID_M_IT = " & h2ID_M_IT & " ORDER BY ID_CFG"
            Call CARGAR_DATAGRIDVIEW_3()
            Call ARMAR_DATAGRIDVIEW_3()
            Call SELECCIONAR_DATAGRIDVIEW_3()
        End If
    End Sub
    Private Sub DGV3_DoubleClick(sender As Object, e As EventArgs) Handles DGV3.DoubleClick
        Call VERIFICAR_ACCESOS("202") : If HAY_ACCESO = False Then : Exit Sub : End If
        If h2ID_M_IT = 0 Then
            MsgBox("Debe seleccionar el registro de la Información del Accidente del que desea Visualizar las Causas Factores y Gravedad", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        If h3ID_CFG = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Accidentes_Laborales_Upd_CFG.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HIGIENE CAUSAS FACTORES Y GRAVEDAD] WHERE ID_M_P = " & h1ID_M_P & " AND ID_M_IT = " & h2ID_M_IT & " ORDER BY ID_CFG"
            Call CARGAR_DATAGRIDVIEW_3()
            Call ARMAR_DATAGRIDVIEW_3()
            Call SELECCIONAR_DATAGRIDVIEW_3()
        End If
    End Sub
    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Call VERIFICAR_ACCESOS("203") : If HAY_ACCESO = False Then : Exit Sub : End If
        If h3ID_CFG = 0 Then
            MsgBox("Debe seleccionar el registro que desea Eliminar", vbInformation, "Mensaje del Sistema")
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
                MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE HIGIENE CAUSAS FACTORES Y GRAVEDAD] WHERE ID_CFG = " & CInt(h3ID_CFG) & ""
                MiSqlCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HIGIENE CAUSAS FACTORES Y GRAVEDAD] WHERE ID_M_P = " & h1ID_M_P & " AND ID_M_IT = " & h2ID_M_IT & " ORDER BY ID_CFG"
            Call CARGAR_DATAGRIDVIEW_3()
            Call ARMAR_DATAGRIDVIEW_3()
        Else
            h3ID_CFG = 0
        End If
    End Sub
    Private Sub DGV4_Click(sender As Object, e As EventArgs) Handles DGV4.Click
        If Me.DGV4.RowCount > 0 Then
            For I = 0 To Me.DGV4.RowCount - 1
                If DGV4.Rows(I).Selected = True Then
                    h4ID_DFA = Me.DGV4.Item(0, I).Value
                    h4ID_M_IT = Me.DGV4.Item(1, I).Value
                    h4ID_M_P = Me.DGV4.Item(2, I).Value
                    h4ID_AMC = Me.DGV4.Item(3, I).Value
                    h4D_AMC = Me.DGV4.Item(4, I).Value
                    h4CONCLUSIONES = Me.DGV4.Item(5, I).Value
                    h4LYN = Me.DGV4.Item(6, I).Value
                    h4NO_REG_ADMI_E = Me.DGV4.Item(7, I).Value
                    h4NO_RES_CLI = Me.DGV4.Item(8, I).Value
                    '-
                    h4RC = Me.DGV4.Item(9, I).Value
                    h4AMC = Me.DGV4.Item(10, I).Value
                    h4AMCMA = Me.DGV4.Item(11, I).Value
                    h4RT = Me.DGV4.Item(12, I).Value
                    h4OBS = Me.DGV4.Item(13, I).Value
                    h4ANEXOS = Me.DGV4.Item(14, I).Value
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Call VERIFICAR_ACCESOS("206") : If HAY_ACCESO = False Then : Exit Sub : End If
        If h4ID_DFA = 0 Then
            MsgBox("Debe seleccionar el registro que desea Eliminar", vbInformation, "Mensaje del Sistema")
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
                MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE HIGIENE DISPOSICION FINAL Y ANEXOS] WHERE ID_DFA = " & CInt(h4ID_DFA) & ""
                MiSqlCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HIGIENE DISPOSICION FINAL Y ANEXOS] WHERE ID_M_P = " & h1ID_M_P & " AND ID_M_IT = " & h2ID_M_IT & " ORDER BY ID_DFA"
            Call CARGAR_DATAGRIDVIEW_4()
            Call ARMAR_DATAGRIDVIEW_4()
        Else
            h4ID_DFA = 0
        End If
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Call VERIFICAR_ACCESOS("204") : If HAY_ACCESO = False Then : Exit Sub : End If
        If h2ID_M_IT = 0 Then
            MsgBox("Debe seleccionar el registro de la Información del Accidente del que desea Visualizar la Disposicion Final y Anexos", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Accidentes_Laborales_Add_DFA.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HIGIENE DISPOSICION FINAL Y ANEXOS] WHERE ID_M_P = " & h1ID_M_P & " AND ID_M_IT = " & h2ID_M_IT & " ORDER BY ID_DFA"
            Call CARGAR_DATAGRIDVIEW_4()
            Call ARMAR_DATAGRIDVIEW_4()
            Call SELECCIONAR_DATAGRIDVIEW_4()
        End If
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Call VERIFICAR_ACCESOS("205") : If HAY_ACCESO = False Then : Exit Sub : End If
        If h2ID_M_IT = 0 Then
            MsgBox("Debe seleccionar el registro de la Información del Accidente del que desea Visualizar la Disposicion Final y Anexos", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        If h4ID_DFA = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Accidentes_Laborales_Upd_DFA.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HIGIENE DISPOSICION FINAL Y ANEXOS] WHERE ID_M_P = " & h1ID_M_P & " AND ID_M_IT = " & h2ID_M_IT & " ORDER BY ID_DFA"
            Call CARGAR_DATAGRIDVIEW_4()
            Call ARMAR_DATAGRIDVIEW_4()
            Call SELECCIONAR_DATAGRIDVIEW_4()
        End If
    End Sub
    Private Sub DGV4_DoubleClick(sender As Object, e As EventArgs) Handles DGV4.DoubleClick
        Call VERIFICAR_ACCESOS("205") : If HAY_ACCESO = False Then : Exit Sub : End If
        If h2ID_M_IT = 0 Then
            MsgBox("Debe seleccionar el registro de la Información del Accidente del que desea Visualizar la Disposicion Final y Anexos", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        If h4ID_DFA = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Accidentes_Laborales_Upd_DFA.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HIGIENE DISPOSICION FINAL Y ANEXOS] WHERE ID_M_P = " & h1ID_M_P & " AND ID_M_IT = " & h2ID_M_IT & " ORDER BY ID_DFA"
            Call CARGAR_DATAGRIDVIEW_4()
            Call ARMAR_DATAGRIDVIEW_4()
            Call SELECCIONAR_DATAGRIDVIEW_4()
        End If
    End Sub
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Call VERIFICAR_ACCESOS("197") : If HAY_ACCESO = False Then : Exit Sub : End If
        If h2ID_M_IT = 0 Then
            MsgBox("Debe seleccionar el registro que desea Eliminar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Esta seguro de Eliminar el registro seleccionado y Todos los registros vinculados?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            Call BORRAR_MAESTRO_DE_HIGIENE_DISPOSICION_FINAL_Y_ANEXOS()
            Call BORRAR_MAESTRO_DE_HIGIENE_CAUSAS_FACTORES_Y_GRAVEDAD()
            Call BORRAR_MAESTRO_DE_HIGIENE_DATOS_DEL_EVENTO()
            Call BORRAR_MAESTRO_DE_HIGIENE_INFORMACION_DEL_ACCIDENTE()
            '-------------------------------------------------------
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HIGIENE INFORMACION DEL ACCIDENTE] WHERE ID_M_P = " & h1ID_M_P & " ORDER BY ID_M_P, NO_PARTE_ACC"
            Call CARGAR_DATAGRIDVIEW_1()
            Call ARMAR_DATAGRIDVIEW_1()
            '-------------------------------------------------------
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HIGIENE DATOS DEL EVENTO] WHERE ID_M_P = " & h1ID_M_P & " AND ID_M_IT = " & h2ID_M_IT & " ORDER BY FECHA_ACC"
            Call CARGAR_DATAGRIDVIEW_2()
            Call ARMAR_DATAGRIDVIEW_2()
            '-------------------------------------------------------
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HIGIENE CAUSAS FACTORES Y GRAVEDAD] WHERE ID_M_P = " & h1ID_M_P & " AND ID_M_IT = " & h2ID_M_IT & " ORDER BY ID_CFG"
            Call CARGAR_DATAGRIDVIEW_3()
            Call ARMAR_DATAGRIDVIEW_3()
            '-------------------------------------------------------
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HIGIENE DISPOSICION FINAL Y ANEXOS] WHERE ID_M_P = " & h1ID_M_P & " AND ID_M_IT = " & h2ID_M_IT & " ORDER BY ID_DFA"
            Call CARGAR_DATAGRIDVIEW_4()
            Call ARMAR_DATAGRIDVIEW_4()
            '-------------------------------------------------------
        End If
    End Sub
    Private Sub BORRAR_MAESTRO_DE_HIGIENE_INFORMACION_DEL_ACCIDENTE()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE HIGIENE INFORMACION DEL ACCIDENTE] WHERE ID_M_IT = " & CInt(h2ID_M_IT) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub BORRAR_MAESTRO_DE_HIGIENE_DATOS_DEL_EVENTO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE HIGIENE DATOS DEL EVENTO] WHERE ID_M_IT = " & CInt(h2ID_M_IT) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub BORRAR_MAESTRO_DE_HIGIENE_CAUSAS_FACTORES_Y_GRAVEDAD()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE HIGIENE CAUSAS FACTORES Y GRAVEDAD] WHERE ID_M_IT = " & CInt(h2ID_M_IT) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub BORRAR_MAESTRO_DE_HIGIENE_DISPOSICION_FINAL_Y_ANEXOS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE HIGIENE DISPOSICION FINAL Y ANEXOS] WHERE ID_M_IT = " & CInt(h2ID_M_IT) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Call VERIFICAR_ACCESOS("190") : If HAY_ACCESO = False Then : Exit Sub : End If
        If h2ID_M_IT = 0 Then
            MsgBox("Debe seleccionar el registro de Información del Accidente que desea Imprimir", vbInformation, "Mensaje del Sistema")
            Me.DGV1.Focus()
            Exit Sub
        End If
        SELECCION = "({MAESTRO_DE_HIGIENE_INFORMACION_DEL_ACCIDENTE.ID_M_IT} = " & h2ID_M_IT & ")"
        PARAMETRO = 16
        Frm_Visor_Reportes.CrystalR.ShowExportButton = True
        Frm_Visor_Reportes.CrystalR.ShowPrintButton = True
        Frm_Visor_Reportes.ShowDialog()
    End Sub
End Class