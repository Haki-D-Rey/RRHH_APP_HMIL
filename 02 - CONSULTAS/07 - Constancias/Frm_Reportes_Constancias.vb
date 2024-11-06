Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Imports Newtonsoft.Json
Public Class Frm_Reportes_Constancias

    Private EstadoCheckedBox As Boolean = False

    Private Sub Frm_Reportes_Constancias_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Reportes_Constancias_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load

        Me.TextBox1.Text = ""
        Me.TextBox6.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.DateTimePicker1.Value = Now
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Me.CheckBox1.Checked = Me.EstadoCheckedBox
        Me.CheckBox2.Checked = Me.EstadoCheckedBox


        Me.TextBox12.Text = "Tnte. Cnel. (Inf.) DEM"
        Me.TextBox11.Text = "JUAN RAMÓN REYES PÉRES"
        Me.TextBox13.Text = "Licenciada"
        Me.TextBox14.Text = "MARIA ISABEL PEREZ MOLINA"
        Me.TextBox8.Text = "Licenciada"
        Me.TextBox16.Text = "MARCELA ESILDA LARIOS MALESPIN"
        Me.TextBox7.Text = "para los tramites que estime convenientes"
        Me.TextBox15.Text = ""
        Me.TextBox1.SelectAll()
        Me.TextBox1.Focus()
    End Sub
    Private Sub Button7_Click(sender As Object, e As System.EventArgs) Handles Button7.Click
        Me.Close()
    End Sub
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
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
    Dim id_EMPLEADO As Integer
    Dim cEMPLEADO As String
    Dim DATO_1, DATO_2 As Boolean
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            cEMPLEADO = CInt(Val(TextBox1.Text))
            TextBox1.Text = Format(CInt(cEMPLEADO), "0000000000")
            If Len(Me.TextBox1.Text) = 0 Or (Me.TextBox1.Text = "0000000000") Then
                MsgBox("El código que se digitó no es válido", vbInformation, "Mensaje del Sistema")
                Me.TextBox1.Focus()
                Exit Sub
            Else
                If CODIGO_REPORTE = "E0009" Or CODIGO_REPORTE = "E0011" Then
                    CADENAsql = "SELECT ID_M_P, CODIGO, APELLIDOS, NOMBRES, N_CEDULA, FEC_ING_PAME, ID_ESTADO_P, ID_N_PROFESIONAL, ID_T_SEXO, PREFIJO FROM [dbo].[VISTA MAESTRO DE PERSONAS] WHERE CODIGO = '" & Me.TextBox1.Text & "' AND ID_ESTADO_P = 1"
                End If
                If CODIGO_REPORTE = "E0010" Or CODIGO_REPORTE = "E0012" Then
                    CADENAsql = "SELECT ID_M_P, CODIGO, APELLIDOS, NOMBRES, N_CEDULA, FEC_ING_PAME, ID_ESTADO_P, ID_N_PROFESIONAL, ID_T_SEXO, PREFIJO FROM [dbo].[VISTA MAESTRO DE PERSONAS] WHERE CODIGO = '" & Me.TextBox1.Text & "' AND ID_ESTADO_P = 2"
                End If
                Call BUSCAR_DATO_1()
                If DATO_1 = True Then
                    If CODIGO_REPORTE = "E0009" Or CODIGO_REPORTE = "E0011" Then
                        Call BUSCAR_DATO_2_ACTIVOS()
                    End If
                    If CODIGO_REPORTE = "E0010" Or CODIGO_REPORTE = "E0012" Then
                        Call BUSCAR_DATO_2_INACTIVOS()
                    End If
                    Call BUSCAR_DATO_3()
                    Call BUSCAR_DATO_5()
                    Call FECHAS()
                    e.Handled = True
                    SendKeys.Send("{TAB}")
                Else
                    MsgBox("El código que se digitó no es válido", vbInformation, "Mensaje del Sistema")
                    Me.TextBox1.SelectAll()
                    Me.TextBox1.Focus()
                    Exit Sub
                End If
            End If
        End If
    End Sub
    Private Sub BUSCAR_DATO_5()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [dbo].[VISTA MAESTRO DE NOMINAS Y SALARIOS] WHERE ID_M_P = " & id_EMPLEADO & " AND ACTIVO = 1", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                Me.TextBox9.Text = MiDataTable.Rows(0).Item(8).ToString
                Me.TextBox9.Text = Format(Val(Me.TextBox9.Text), "#,##0.00")
                Me.TextBox10.Text = LETRAS_CONSTANCIAS(Me.TextBox9.Text)
                Me.TextBox10.Text = "(" & Me.TextBox10.Text & ")"
            Else
                Me.TextBox9.Text = "0.0"
                Me.TextBox10.Text = "."
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub BUSCAR_DATO_3()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [dbo].[FIRMANTE DE CONSTANCIAS] WHERE ID_FIRMA= 1", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                Me.TextBox12.Text = MiDataTable.Rows(0).Item(1).ToString
                Me.TextBox11.Text = MiDataTable.Rows(0).Item(2).ToString
            Else
                Me.TextBox12.Text = ""
                Me.TextBox1.Text = ""
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim n1, n2, n3, n4, n5, n6, n7, n8 As String
    Private Sub BUSCAR_DATO_2_INACTIVOS()
        DE = ""
        SE = ""
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT CODIGO, FEC_MOV, D_N1, D_N2, D_N3, D_N4, D_N5, D_N6, D_N7, D_N8, D_CARGO_ES, ID_T_M_EXP FROM [dbo].[VISTA HISTORICO DE MOVIMIENTO] WHERE CODIGO = '" & Me.TextBox1.Text & "' AND ID_T_M_EXP = 1 ORDER BY FEC_MOV DESC", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                Me.DateTimePicker2.Value = MiDataTable.Rows(0).Item(1).ToString
                '---------------------------------------------------------------------------------- SERVICIO
                If MiDataTable.Rows(0).Item(9).ToString <> "" Then  'NIVEL 8
                    SE = MiDataTable.Rows(0).Item(9).ToString
                    GoTo SERVICIO
                Else
                    If MiDataTable.Rows(0).Item(8).ToString <> "" Then  'NIVEL 7
                        SE = MiDataTable.Rows(0).Item(8).ToString
                        GoTo SERVICIO
                    Else
                        If MiDataTable.Rows(0).Item(7).ToString <> "" Then   'NIVEL 6
                            SE = MiDataTable.Rows(0).Item(7).ToString
                            GoTo SERVICIO
                        Else
                            If MiDataTable.Rows(0).Item(6).ToString <> "" Then  'NIVEL 5
                                SE = MiDataTable.Rows(0).Item(6).ToString
                                GoTo SERVICIO
                            Else
                                If MiDataTable.Rows(0).Item(5).ToString <> "" Then  'NIVEL 4
                                    SE = MiDataTable.Rows(0).Item(5).ToString
                                    GoTo SERVICIO
                                Else
                                    If MiDataTable.Rows(0).Item(4).ToString <> "" Then  'NIVEL 3
                                        SE = MiDataTable.Rows(0).Item(4).ToString
                                        GoTo SERVICIO
                                    Else
                                        If MiDataTable.Rows(0).Item(3).ToString <> "" Then  'NIVEL 2
                                            SE = MiDataTable.Rows(0).Item(3).ToString
                                            GoTo SERVICIO
                                        Else
                                            If MiDataTable.Rows(0).Item(2).ToString <> "" Then  'NIVEL 1
                                                SE = MiDataTable.Rows(0).Item(2).ToString
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
                If MiDataTable.Rows(0).Item(9).ToString <> "" Then  'NIVEL 8
                    DE = MiDataTable.Rows(0).Item(9).ToString
                    If DE <> SE Then
                        GoTo DEPARTAMENTO
                    Else
                        GoTo LEE_1
                    End If
                Else
LEE_1:
                    If MiDataTable.Rows(0).Item(8).ToString <> "" Then  'NIVEL 7
                        DE = MiDataTable.Rows(0).Item(8).ToString
                        If DE <> SE Then
                            GoTo DEPARTAMENTO
                        Else
                            GoTo LEE_2
                        End If
                    Else
LEE_2:
                        If MiDataTable.Rows(0).Item(7).ToString <> "" Then  'NIVEL 6
                            DE = MiDataTable.Rows(0).Item(7).ToString
                            If DE <> SE Then
                                GoTo DEPARTAMENTO
                            Else
                                GoTo LEE_3
                            End If
                        Else
LEE_3:
                            If MiDataTable.Rows(0).Item(6).ToString <> "" Then  'NIVEL 5
                                DE = MiDataTable.Rows(0).Item(6).ToString
                                If DE <> SE Then
                                    GoTo DEPARTAMENTO
                                Else
                                    GoTo LEE_4
                                End If
                            Else
LEE_4:
                                If MiDataTable.Rows(0).Item(5).ToString <> "" Then  'NIVEL 4
                                    DE = MiDataTable.Rows(0).Item(5).ToString
                                    If DE <> SE Then
                                        GoTo DEPARTAMENTO
                                    Else
                                        GoTo LEE_5
                                    End If
                                Else
LEE_5:
                                    If MiDataTable.Rows(0).Item(4).ToString <> "" Then  'NIVEL 3
                                        DE = MiDataTable.Rows(0).Item(4).ToString
                                        If DE <> SE Then
                                            GoTo DEPARTAMENTO
                                        Else
                                            GoTo LEE_6
                                        End If
                                    Else
LEE_6:
                                        If MiDataTable.Rows(0).Item(3).ToString <> "" Then  'NIVEL 2
                                            DE = MiDataTable.Rows(0).Item(3).ToString
                                            If DE <> SE Then
                                                GoTo DEPARTAMENTO
                                            Else
                                                GoTo LEE_7
                                            End If
                                        Else
LEE_7:
                                            If MiDataTable.Rows(0).Item(2).ToString <> "" Then  'NIVEL 1
                                                DE = MiDataTable.Rows(0).Item(2).ToString
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
                If SE = "" Then
                    Me.TextBox5.Text = DE
                Else
                    Me.TextBox5.Text = DE & "\" & SE
                End If
                Me.TextBox4.Text = MiDataTable.Rows(0).Item(10).ToString
                DATO_2 = True
            Else
                Me.DateTimePicker2.Value = Mid(Now, 1, 10)
                Me.TextBox5.Text = ""
                Me.TextBox4.Text = ""
                DATO_2 = False
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
    Private Sub BUSCAR_DATO_2_ACTIVOS()
        DE = ""
        SE = ""
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT CODIGO, D_N1, D_N2, D_N3, D_N4, D_N5, D_N6, D_N7, D_N8, D_CARGO_ES, ID_SITUACION FROM [dbo].[VISTA MAESTRO DE CARGOS SIT] WHERE CODIGO = '" & Me.TextBox1.Text & "' AND ID_SITUACION = 1", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
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
                If SE = "" Then
                    Me.TextBox5.Text = DE
                Else
                    Me.TextBox5.Text = DE & "\" & SE
                End If
                Me.TextBox4.Text = MiDataTable.Rows(0).Item(9).ToString
                DATO_2 = True
            Else
                Me.TextBox5.Text = ""
                Me.TextBox4.Text = ""
                DATO_2 = False
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim NP As Integer
    Dim SEX As String
    Private Sub BUSCAR_DATO_1()
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
                id_EMPLEADO = MiDataTable.Rows(0).Item(0).ToString
                Me.TextBox2.Text = MiDataTable.Rows(0).Item(3).ToString & " " & MiDataTable.Rows(0).Item(2).ToString
                Me.TextBox3.Text = MiDataTable.Rows(0).Item(4).ToString
                Me.DateTimePicker1.Value = MiDataTable.Rows(0).Item(5).ToString
                If MiDataTable.Rows(0).Item(8).ToString = 1 Then    'FEMENINO
                    NP = MiDataTable.Rows(0).Item(7).ToString
                    SEX = "la "
                End If
                If MiDataTable.Rows(0).Item(8).ToString = 2 Then    'MASCULINO
                    NP = MiDataTable.Rows(0).Item(7).ToString
                    SEX = "el "
                End If
                If MiDataTable.Rows(0).Item(8).ToString = 3 Then    '.
                    NP = MiDataTable.Rows(0).Item(7).ToString
                    SEX = " "
                End If
                Me.TextBox6.Text = MiDataTable.Rows(0).Item(9).ToString
                DATO_1 = True
            Else
                Me.TextBox2.Text = ""
                Me.TextBox3.Text = ""
                Me.DateTimePicker1.Value = Now
                NP = 0
                Me.TextBox6.Text = ""
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
    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
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
    Private Sub DateTimePicker1_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker1.KeyDown
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
    Private Sub CheckBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox1.KeyDown
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
    Private Sub TextBox14_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox14.KeyDown
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
    Private Sub TextBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox7.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub TextBox16_TextChanged(sender As Object, e As EventArgs) Handles TextBox16.TextChanged

    End Sub

    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs) Handles TextBox14.TextChanged

    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged

    End Sub

    Private Sub TextBox13_TextChanged(sender As Object, e As EventArgs) Handles TextBox13.TextChanged

    End Sub

    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged

    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox15_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox15.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Dim F As Date
    Private Sub FECHAS()
        Dim AÑO_LETRAS As String
        Dim DIAS_LETRAS As String
        Dim FECHA_DIA_LETRAS As String
        Dim FECHA_DIA As String
        Dim FECHA_MES As String
        FECHA_DIA_LETRAS = ""
        FECHA_DIA = ""
        FECHA_MES = ""
        F = Now
        FECHA_DIA_LETRAS = Now.ToString("dddd")
        FECHA_DIA = Now.Day
        FECHA_MES = Month(Now).ToString
        If FECHA_MES = "1" Then : FECHA_MES = "Enero" : End If
        If FECHA_MES = "2" Then : FECHA_MES = "Febrero" : End If
        If FECHA_MES = "3" Then : FECHA_MES = "Marzo" : End If
        If FECHA_MES = "4" Then : FECHA_MES = "Abril" : End If
        If FECHA_MES = "5" Then : FECHA_MES = "Mayo" : End If
        If FECHA_MES = "6" Then : FECHA_MES = "Junio" : End If
        If FECHA_MES = "7" Then : FECHA_MES = "Julio" : End If
        If FECHA_MES = "8" Then : FECHA_MES = "Agosto" : End If
        If FECHA_MES = "9" Then : FECHA_MES = "Septiembre" : End If
        If FECHA_MES = "10" Then : FECHA_MES = "Octubre" : End If
        If FECHA_MES = "11" Then : FECHA_MES = "Noviembre" : End If
        If FECHA_MES = "12" Then : FECHA_MES = "Diciembre" : End If

        DIAS_LETRAS = LETRAS(FECHA_DIA)
        AÑO_LETRAS = LETRAS(Now.Year)

        If FECHA_DIA = "1" Then
            Me.TextBox15.Text = "al primer dia del mes de " & FECHA_MES & " del año " & AÑO_LETRAS
        Else
            Me.TextBox15.Text = "a los " & DIAS_LETRAS & " dias del mes de " & FECHA_MES & " del año " & AÑO_LETRAS
        End If
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        VALOR01 = Me.TextBox1.Text      'CODIGO
        VALOR02 = SEX & Me.TextBox6.Text      'PRONOMBRE
        VALOR03 = Me.TextBox2.Text      'NOMBRES Y APELLIDOS
        VALOR04 = Me.TextBox3.Text      'CEDULA
        VALOR05 = Mid(Me.DateTimePicker1.Value, 1, 10)    'FECHA INGRESO PAME
        VALOR06 = Me.TextBox4.Text      'CARGO
        VALOR07 = Me.TextBox5.Text      'DEPARTAMENTO
        If Me.CheckBox1.Checked = True Then
            VALOR09 = Me.TextBox12.Text 'GRADO JEFE
            VALOR10 = Me.TextBox11.Text 'NOMBRES Y APELLIDOS JEFE
            VALOR15 = Me.CheckBox1.Text 'CARGO JEFE 
            VALOR17 = $"El suscrito {Me.CheckBox1.Text} del Hospital Militar, hace constar:"

        ElseIf Me.CheckBox2.Checked = True Then
            VALOR09 = Me.TextBox13.Text 'GRADO 1ER OFICIAL
            VALOR10 = Me.TextBox14.Text 'NOMBRES Y APELLIDOS 1ER OFICIAL
            VALOR15 = Me.CheckBox2.Text 'CARGO 1ER OFICIAL
            VALOR17 = $"El suscrito {Me.CheckBox2.Text} del Hospital Militar, hace constar:"
        Else
            VALOR09 = Me.TextBox8.Text 'GRADO JEFE
            VALOR10 = Me.TextBox16.Text 'NOMBRES Y APELLIDOS JEFE
            VALOR15 = Me.CheckBox3.Text 'CARGO JEFE

            VALOR17 = $"La suscrita {Me.CheckBox3.Text} del Hospital Militar, hace constar:"
        End If
        VALOR11 = Me.TextBox7.Text      'ASUNTO
        VALOR12 = Me.TextBox15.Text     'FECHA EN LETRAS
        VALOR13 = Mid(Me.DateTimePicker2.Value, 1, 10)    'FECHA BAJA
        VALOR14 = Me.TextBox9.Text     'SALARIO
        VALOR16 = Me.TextBox10.Text    'SALARIO LETRAS
        If CODIGO_REPORTE = "E0009" Then                    'Personal Activo
            PARAMETRO = 130
        End If
        If CODIGO_REPORTE = "E0010" Then                    'Personal Inactivo
            PARAMETRO = 131
        End If
        If CODIGO_REPORTE = "E0011" Then                    'Salarial
            PARAMETRO = 132
        End If
        If CODIGO_REPORTE = "E0012" Then                    'Solvencia
            PARAMETRO = 133
        End If
        SELECCION = ""
        Frm_Visor_Reportes.ShowDialog()
    End Sub
    Private Sub TextBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox9.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            Me.TextBox9.Text = Format(Val(Me.TextBox9.Text), "#,##0.00")
            Me.TextBox10.Text = LETRAS_CONSTANCIAS(Me.TextBox9.Text)
            Me.TextBox10.Text = "(" & Me.TextBox10.Text & ")"
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox1_DoubleClick(sender As Object, e As EventArgs) Handles TextBox1.DoubleClick
        'Frm_Expedientes_Buscar_Activos.ShowDialog()
        Frm_Seleccion_Consultas_y_Reportes_Buscar.ShowDialog()
        If CERRAR = False Then
            Me.TextBox1.Text = bCODIGO
            Me.TextBox1.SelectAll()
            Me.TextBox1.Focus()
        End If
        'Frm_Reportes_Buscar.ShowDialog()
        'Me.TextBox1.Text = CE
    End Sub
    Private Sub TextBox9_DoubleClick(sender As Object, e As EventArgs) Handles TextBox9.DoubleClick
        Frm_Reportes_Constancias_Salarios.ShowDialog()
        Me.TextBox9.Text = Format(Val(TOTAL_SALARIOS), "#,##0.00")
        Me.TextBox10.Text = LETRAS_CONSTANCIAS(Me.TextBox9.Text)
        Me.TextBox10.Text = "(" & Me.TextBox10.Text & ")"
    End Sub

    ' Variable para rastrear el último checkbox clickeado
    Private ultimoCheckBoxClickeado As CheckBox

    ' Manejador de evento para todos los checkboxes
    Private Sub CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged, CheckBox2.CheckedChanged, CheckBox3.CheckedChanged
        Dim clickedCheckBox As CheckBox = DirectCast(sender, CheckBox)

        ' Verifica si el checkbox clickeado es diferente al último clickeado
        If clickedCheckBox IsNot ultimoCheckBoxClickeado Then
            ' Desactiva todos los checkboxes excepto el clickeado
            For Each checkBox As CheckBox In {CheckBox1, CheckBox2, CheckBox3}
                If checkBox IsNot clickedCheckBox Then
                    checkBox.Checked = False
                End If
            Next
        End If

        ' Actualiza el último checkbox clickeado
        ultimoCheckBoxClickeado = clickedCheckBox
    End Sub

End Class