Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Public Class Frm_Situacion_Personal
    Dim FECHA1, FECHA2 As String
    Private Sub Frm_Situacion_Personal_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            SITUACIONES = False
            Me.TextBox2.SelectAll()
            Me.TextBox2.Focus()
            Frm_Principal.Timer1.Start()
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Situacion_Personal_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Frm_Principal.Timer1.Stop()
        RemoveHandler Me.DateTimePicker1.ValueChanged, AddressOf DateTimePicker1_ValueChanged
        RemoveHandler Me.DateTimePicker2.ValueChanged, AddressOf DateTimePicker2_ValueChanged
        SITUACIONES = True
        id_EMPLEADO = 0
        Call Clases.BUSCAR_FECHA_Y_HORA()
        Me.DateTimePicker1.Value = "01/" & FECHA_SERVIDOR.Month & "/" & FECHA_SERVIDOR.Year
        Me.DateTimePicker2.Value = FECHA_SERVIDOR.Day & "/" & FECHA_SERVIDOR.Month & "/" & FECHA_SERVIDOR.Year
        FECHA01 = Me.DateTimePicker1.Text
        Me.DateTimePicker10.Value = Me.DateTimePicker1.Text
        FECHA1 = Mid(FECHA01, 1, 10)
        FECHA02 = Me.DateTimePicker2.Text
        Me.DateTimePicker20.Value = Me.DateTimePicker2.Text
        Me.DateTimePicker30.Value = Me.DateTimePicker1.Text
        FECHA2 = Mid(FECHA02, 1, 10)
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""

        If isuID_TIPO_USUARIO = 5 Then
            CADENAsql = "SET LANGUAGE ESPAÑOL Select ID, NIVEL1, NIVEL2, NIVEL3, NIVEL4, NIVEL5, ORD, ID_M_A, ID_M_P, CODIGO, UPPER(DATENAME (dw, DIA)) AS DIALETRAS, DIA, ID_T_SIT_P, D_T_SIT_P, ID_SIT_P, D_SIT_P, SIGLA, VALOR_SIT, VALOR_DIA, OBSERVACIONES, DETALLE_1, DETALLE_2, DETALLE_3, D_ADMIN, USUARIO_ACT, FECHA_ACT from [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE CODIGO = '0000000000' AND DIA BETWEEN '" & FECHA1 & "' AND '" & FECHA2 & "' AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY CODIGO, DIA SET LANGUAGE ENGLISH"
        Else
            CADENAsql = "SET LANGUAGE ESPAÑOL Select ID, NIVEL1, NIVEL2, NIVEL3, NIVEL4, NIVEL5, ORD, ID_M_A, ID_M_P, CODIGO, UPPER(DATENAME (dw, DIA)) AS DIALETRAS, DIA, ID_T_SIT_P, D_T_SIT_P, ID_SIT_P, D_SIT_P, SIGLA, VALOR_SIT, VALOR_DIA, OBSERVACIONES, DETALLE_1, DETALLE_2, DETALLE_3, D_ADMIN, USUARIO_ACT, FECHA_ACT from [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE CODIGO = '0000000000' AND DIA BETWEEN '" & FECHA1 & "' AND '" & FECHA2 & "' ORDER BY CODIGO, DIA SET LANGUAGE ENGLISH"
        End If

        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
        '--------------------------
        Call DIAS_SUBSIDIOS()
        Call DIAS_AI()
        Call DIAS_V()
        Call DIAS_OM()
        Label11.Text = vOM + vAI
        '--------------------------
        Call CARGAR_DATAGRIDVIEW_15()
        Call ARMAR_DATAGRIDVIEW_15()
        Me.DGV.DataSource = Nothing
        Me.DGV15.DataSource = Nothing
        AddHandler Me.DateTimePicker1.ValueChanged, AddressOf DateTimePicker1_ValueChanged
        AddHandler Me.DateTimePicker2.ValueChanged, AddressOf DateTimePicker2_ValueChanged
        Me.TextBox2.SelectAll()
        Me.TextBox2.Focus()
    End Sub
    Private Sub Button7_Click(sender As Object, e As System.EventArgs) Handles Button7.Click
        SITUACIONES = False
        Me.TextBox2.SelectAll()
        Me.TextBox2.Focus()
        Frm_Principal.Timer1.Start()
        Me.Close()
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_15()
        '-------------*************----------------------------------------
        FECHA01 = "Convert(VARCHAR(10), Convert(Date, '" + msDIA + "', 105), 23)"
        '-------------*************----------------------------------------
        CADENAsql = "SET LANGUAGE Español SELECT [IDEMPRESA], [RELOJ], [ID_EMPLEADO], DATENAME (dw, DIA) AS DIAL, [FECHA_MARCA], [HORA_MARCA], [FECHA_CARGA], [ARCHIVO], [NUMERO], [PROGRAMADO], [TIPO_MARCA], [OBSERVACION] FROM [dbo].[VISTA MAESTRO DE MARCAS] WHERE (ID_EMPLEADO = '" & msCODIGO & "') AND (FECHA_MARCA = " & FECHA01 & ")  ORDER BY FECHA_MARCA, HORA_MARCA SET LANGUAGE ENGLISH"

        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE MARCAS]")
        DGV15.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE MARCAS]")
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_15()
        Me.DGV15.Columns(0).HeaderText = "IDEMPRESA"
        Me.DGV15.Columns(0).Width = 0
        Me.DGV15.Columns(0).Visible = False

        Me.DGV15.Columns(1).HeaderText = "Ubicacion Reloj"
        Me.DGV15.Columns(1).Width = 200
        Me.DGV15.Columns(1).Visible = True
        Me.DGV15.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV15.Columns(2).HeaderText = "ID_EMPLEADO"
        Me.DGV15.Columns(2).Width = 0
        Me.DGV15.Columns(2).Visible = False
        Me.DGV15.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV15.Columns(3).HeaderText = "Dia"
        Me.DGV15.Columns(3).Width = 0
        Me.DGV15.Columns(3).Visible = False
        Me.DGV15.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV15.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV15.Columns(4).HeaderText = "Fecha"
        Me.DGV15.Columns(4).Width = 0
        Me.DGV15.Columns(4).Visible = False
        Me.DGV15.Columns(4).DefaultCellStyle.Format = "d"
        Me.DGV15.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV15.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV15.Columns(5).HeaderText = "Hora Marca"
        Me.DGV15.Columns(5).Width = 200
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

        Me.DGV15.Columns(10).HeaderText = "Tipo de Marca"
        Me.DGV15.Columns(10).Width = 200
        Me.DGV15.Columns(10).Visible = True
        Me.DGV15.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV15.Columns(11).HeaderText = "Observaciones"
        Me.DGV15.Columns(11).Width = 510
        Me.DGV15.Columns(11).Visible = True
        Me.DGV15.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE SITUACION DE PERSONAL]")
        DGV.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE SITUACION DE PERSONAL]")
        Me.Label2.Text = DGV.RowCount & " Registro(s) Existente(s) ---> Seleccionados: " & DGV.SelectedRows.Count
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW()
        Me.DGV.Columns(0).HeaderText = "Id"
        Me.DGV.Columns(0).Width = 0
        Me.DGV.Columns(0).Visible = False
        Me.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(1).HeaderText = "Nivel 1"
        Me.DGV.Columns(1).Width = 0
        Me.DGV.Columns(1).Visible = False
        Me.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(2).HeaderText = "Nivel 2"
        Me.DGV.Columns(2).Width = 0
        Me.DGV.Columns(2).Visible = False
        Me.DGV.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(3).HeaderText = "Nivel 3"
        Me.DGV.Columns(3).Width = 0
        Me.DGV.Columns(3).Visible = False
        Me.DGV.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(4).HeaderText = "Nivel 4"
        Me.DGV.Columns(4).Width = 0
        Me.DGV.Columns(4).Visible = False
        Me.DGV.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(5).HeaderText = "Nivel 5"
        Me.DGV.Columns(5).Width = 0
        Me.DGV.Columns(5).Visible = False
        Me.DGV.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(6).HeaderText = "ORD"
        Me.DGV.Columns(6).Width = 0
        Me.DGV.Columns(6).Visible = False

        Me.DGV.Columns(7).HeaderText = "Id"
        Me.DGV.Columns(7).Width = 10
        Me.DGV.Columns(7).Visible = True
        Me.DGV.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(7).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(8).HeaderText = "ID_M_P"
        Me.DGV.Columns(8).Width = 0
        Me.DGV.Columns(8).Visible = False

        Me.DGV.Columns(9).HeaderText = "CODIGO"
        Me.DGV.Columns(9).Width = 0
        Me.DGV.Columns(9).Visible = False

        Me.DGV.Columns(10).HeaderText = "Dia"
        Me.DGV.Columns(10).Width = 80
        Me.DGV.Columns(10).Visible = True
        Me.DGV.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(11).HeaderText = "Fecha"
        Me.DGV.Columns(11).Width = 70
        Me.DGV.Columns(11).Visible = True
        Me.DGV.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(11).DefaultCellStyle.BackColor = Color.Khaki

        Me.DGV.Columns(12).HeaderText = "ID_T_SIT_P"
        Me.DGV.Columns(12).Width = 0
        Me.DGV.Columns(12).Visible = False

        Me.DGV.Columns(13).HeaderText = "Tipo Situacion"
        Me.DGV.Columns(13).Width = 90
        Me.DGV.Columns(13).Visible = True
        Me.DGV.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(14).HeaderText = "ID_T_SIT"
        Me.DGV.Columns(14).Width = 0
        Me.DGV.Columns(14).Visible = False

        Me.DGV.Columns(15).HeaderText = "Situacion"
        Me.DGV.Columns(15).Width = 90
        Me.DGV.Columns(15).Visible = True
        Me.DGV.Columns(15).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(15).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DGV.Columns(15).DefaultCellStyle.BackColor = Color.LightBlue

        Me.DGV.Columns(16).HeaderText = "Sigla"
        Me.DGV.Columns(16).Width = 0
        Me.DGV.Columns(16).Visible = False
        Me.DGV.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(16).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(17).HeaderText = "Valor"
        Me.DGV.Columns(17).Width = 50
        Me.DGV.Columns(17).Visible = True
        Me.DGV.Columns(17).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(17).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Me.DGV.Columns(18).HeaderText = "Acumulado x Dia"
        Me.DGV.Columns(18).Width = 0
        Me.DGV.Columns(18).Visible = False
        Me.DGV.Columns(18).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(18).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Me.DGV.Columns(19).HeaderText = "Observaciones"
        Me.DGV.Columns(19).Width = 180
        Me.DGV.Columns(19).Visible = True
        Me.DGV.Columns(19).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(19).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(20).HeaderText = "Datos Generales (Subsidio)"
        Me.DGV.Columns(20).Width = 180
        Me.DGV.Columns(20).Visible = True
        Me.DGV.Columns(20).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(20).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(21).HeaderText = "Medico (Subsidio)"
        Me.DGV.Columns(21).Width = 180
        Me.DGV.Columns(21).Visible = True
        Me.DGV.Columns(21).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(21).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(22).HeaderText = "Diagnostico Medico (Subsidio)"
        Me.DGV.Columns(22).Width = 180
        Me.DGV.Columns(22).Visible = True
        Me.DGV.Columns(22).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(22).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(23).HeaderText = "D_ADMIN"
        Me.DGV.Columns(23).Width = 0
        Me.DGV.Columns(23).Visible = False

        Me.DGV.Columns(24).HeaderText = "USUARIO_ACT"
        Me.DGV.Columns(24).Width = 0
        Me.DGV.Columns(24).Visible = False

        Me.DGV.Columns(25).HeaderText = "FECHA_ACT"
        Me.DGV.Columns(25).Width = 0
        Me.DGV.Columns(25).Visible = False
    End Sub
    Private Sub Frm_Situacion_Personal_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        SITUACIONES = False
    End Sub
    Private Sub Frm_Situacion_Personal_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        SITUACIONES = False
    End Sub
    Private Sub Frm_Situacion_Personal_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        SITUACIONES = False
    End Sub
    Private Sub Frm_Situacion_Personal_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        SITUACIONES = False
    End Sub
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call VERIFICAR_ACCESOS("184") : If HAY_ACCESO = False Then : Exit Sub : End If
            cEMPLEADO = CInt(Val(TextBox2.Text))
            TextBox2.Text = Format(CInt(cEMPLEADO), "0000000000")
            If Len(Me.TextBox2.Text) = 0 Or (Me.TextBox2.Text = "0000000000") Then
                MsgBox("El código que se digitó no es válido", vbInformation, "Mensaje del Sistema")
                Me.TextBox2.Focus()
                Exit Sub
            Else
                If isuID_TIPO_USUARIO = 5 Then
                    CADENAsql = "SELECT ID_M_P, CODIGO, APELLIDOS, NOMBRES, N_CEDULA, FEC_ING_PAME, ID_ESTADO_P, ID_N_PROFESIONAL, D_T_SEXO FROM [dbo].[VISTA MAESTRO DE PERSONAS] WHERE CODIGO = '" & Me.TextBox2.Text & "' AND ID_ESTADO_P = 1 AND D_ADMIN = '" & isuUSUARIO & "'"
                Else
                    CADENAsql = "SELECT ID_M_P, CODIGO, APELLIDOS, NOMBRES, N_CEDULA, FEC_ING_PAME, ID_ESTADO_P, ID_N_PROFESIONAL, D_T_SEXO FROM [dbo].[VISTA MAESTRO DE PERSONAS] WHERE CODIGO = '" & Me.TextBox2.Text & "' AND ID_ESTADO_P = 1"
                End If
                Call BUSCAR_DATO_1()
                If DATO_1 = True Then
                    FECHA01 = Mid(Me.DateTimePicker1.Text, 1, 10)
                    FECHA1 = Mid(FECHA01, 1, 10)
                    FECHA02 = Mid(Me.DateTimePicker2.Text, 1, 10)
                    FECHA2 = Mid(FECHA02, 1, 10)
                    If isuID_TIPO_USUARIO = 5 Then
                        CADENAsql = "SET LANGUAGE ESPAÑOL Select ID, NIVEL1, NIVEL2, NIVEL3, NIVEL4, NIVEL5, ORD, ID_M_A, ID_M_P, CODIGO, UPPER(DATENAME (dw, DIA)) AS DIALETRAS, DIA, ID_T_SIT_P, D_T_SIT_P, ID_SIT_P, D_SIT_P, SIGLA, VALOR_SIT, VALOR_DIA, OBSERVACIONES, DETALLE_1, DETALLE_2, DETALLE_3, D_ADMIN, USUARIO_ACT, FECHA_ACT from [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE CODIGO = '" & Me.TextBox2.Text & "' AND DIA BETWEEN '" & FECHA1 & "' AND '" & FECHA2 & "' AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY CODIGO, DIA SET LANGUAGE ENGLISH"
                    Else
                        CADENAsql = "SET LANGUAGE ESPAÑOL Select ID, NIVEL1, NIVEL2, NIVEL3, NIVEL4, NIVEL5, ORD, ID_M_A, ID_M_P, CODIGO, UPPER(DATENAME (dw, DIA)) AS DIALETRAS, DIA, ID_T_SIT_P, D_T_SIT_P, ID_SIT_P, D_SIT_P, SIGLA, VALOR_SIT, VALOR_DIA, OBSERVACIONES, DETALLE_1, DETALLE_2, DETALLE_3, D_ADMIN, USUARIO_ACT, FECHA_ACT from [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE CODIGO = '" & Me.TextBox2.Text & "' AND DIA BETWEEN '" & FECHA1 & "' AND '" & FECHA2 & "' ORDER BY CODIGO, DIA SET LANGUAGE ENGLISH"
                    End If
                    Call CARGAR_DATAGRIDVIEW()
                    Call ARMAR_DATAGRIDVIEW()
                    '--------------------------
                    Call DIAS_SUBSIDIOS()
                    Call DIAS_AI()
                    Call DIAS_V()
                    Call DIAS_OM()
                    Label11.Text = vOM + vAI
                    '--------------------------
                    Call CARGAR_DATAGRIDVIEW_15()
                    Call ARMAR_DATAGRIDVIEW_15()
                    Call BUSCAR_CARGO()
                Else
                    MsgBox("El código que se digitó no es válido o no se encuentra Asignado al usuario Actual", vbInformation, "Mensaje del Sistema")
                    Me.TextBox2.SelectAll()
                    Me.TextBox2.Focus()
                    Exit Sub
                End If
            End If
        End If
    End Sub
    Private Sub BUSCAR_CARGO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            CADENAsql = "SELECT ID_M_C, CODIGO, D_CARGO_ES, ID_SITUACION FROM [dbo].[VISTA MAESTRO DE CARGOS SIT] WHERE CODIGO = '" & Me.TextBox2.Text & "' AND ID_SITUACION = 1"
            MiDataAdapter = New SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                Me.TextBox4.Text = MiDataTable.Rows(0).Item(2).ToString
            Else
                Me.TextBox4.Text = ""
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
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
            MiDataAdapter = New SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                id_EMPLEADO = MiDataTable.Rows(0).Item(0).ToString
                Me.TextBox1.Text = MiDataTable.Rows(0).Item(3).ToString & " " & MiDataTable.Rows(0).Item(2).ToString
                Me.TextBox2.Text = MiDataTable.Rows(0).Item(1).ToString
                Me.TextBox3.Text = Mid(MiDataTable.Rows(0).Item(5).ToString, 1, 10)
                'Me.DateTimePicker1.Value = Mid(MiDataTable.Rows(0).Item(5).ToString, 1, 10)
                Call Clases.BUSCAR_FECHA_Y_HORA()
                Me.DateTimePicker1.Value = "01/" & FECHA_SERVIDOR.Month & "/" & FECHA_SERVIDOR.Year
                Me.TextBox5.Text = MiDataTable.Rows(0).Item(8).ToString
                DATO_1 = True
            Else
                id_EMPLEADO = 0
                Me.TextBox1.Text = ""
                Me.TextBox2.Text = ""
                Me.TextBox3.Text = ""
                Me.TextBox4.Text = ""
                Me.DateTimePicker1.Value = Mid(Now, 1, 10)
                Me.TextBox5.Text = ""
                DATO_1 = False
            End If
            Me.DateTimePicker10.Value = Me.DateTimePicker1.Value
            Me.DateTimePicker20.Value = Me.DateTimePicker2.Value
            Me.DateTimePicker30.Value = Me.DateTimePicker1.Value
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim ultima_fecha As String
    Dim UF As Boolean
    Private Sub BUSCAR_ULTIMA_FECHA_ACTUALIZADA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable1 As New DataTable
        Dim MiDataAdapter1 As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter1 = New SqlDataAdapter("SELECT * FROM [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE [CODIGO] = '" & Me.TextBox2.Text & "' ORDER BY [DIA] DESC", Conexion.CxRRHH)
            MiDataTable1.Clear()
            MiDataAdapter1.Fill(MiDataTable1)
            If MiDataTable1.Rows.Count > 0 Then
                ultima_fecha = Mid(MiDataTable1.Rows(0).Item(7).ToString, 1, 10)    'FECHA DE INGRESO/ULTIMA FECHA ACTUALIZADA
                UF = True
            Else
                UF = False
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Me.DateTimePicker10.Value = Me.DateTimePicker1.Value
        Me.DateTimePicker30.Value = Me.DateTimePicker1.Value
        FECHA01 = Me.DateTimePicker1.Text
        FECHA1 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA01 + "', 105), 23)"
        If Len(Me.TextBox2.Text) = 0 Or (Me.TextBox2.Text = "0000000000") Then
            Exit Sub
        End If
        FECHA1 = Mid(FECHA01, 1, 10)
        If isuID_TIPO_USUARIO = 5 Then
            CADENAsql = "SET LANGUAGE ESPAÑOL Select ID, NIVEL1, NIVEL2, NIVEL3, NIVEL4, NIVEL5, ORD, ID_M_A, ID_M_P, CODIGO, UPPER(DATENAME (dw, DIA)) AS DIALETRAS, DIA, ID_T_SIT_P, D_T_SIT_P, ID_SIT_P, D_SIT_P, SIGLA, VALOR_SIT, VALOR_DIA, OBSERVACIONES, DETALLE_1, DETALLE_2, DETALLE_3, D_ADMIN, USUARIO_ACT, FECHA_ACT from [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE CODIGO = '" & Me.TextBox2.Text & "' AND DIA BETWEEN '" & FECHA1 & "' AND '" & FECHA2 & "' AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY CODIGO, DIA SET LANGUAGE ENGLISH"
        Else
            CADENAsql = "SET LANGUAGE ESPAÑOL Select ID, NIVEL1, NIVEL2, NIVEL3, NIVEL4, NIVEL5, ORD, ID_M_A, ID_M_P, CODIGO, UPPER(DATENAME (dw, DIA)) AS DIALETRAS, DIA, ID_T_SIT_P, D_T_SIT_P, ID_SIT_P, D_SIT_P, SIGLA, VALOR_SIT, VALOR_DIA, OBSERVACIONES, DETALLE_1, DETALLE_2, DETALLE_3, D_ADMIN, USUARIO_ACT, FECHA_ACT from [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE CODIGO = '" & Me.TextBox2.Text & "' AND DIA BETWEEN '" & FECHA1 & "' AND '" & FECHA2 & "' ORDER BY CODIGO, DIA SET LANGUAGE ENGLISH"
        End If
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
        '--------------------------
        Call DIAS_SUBSIDIOS()
        Call DIAS_AI()
        Call DIAS_V()
        Call DIAS_OM()
        Label11.Text = vOM + vAI
        '--------------------------
        Call CARGAR_DATAGRIDVIEW_15()
        Call ARMAR_DATAGRIDVIEW_15()
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Call VERIFICAR_ACCESOS("187") : If HAY_ACCESO = False Then : Exit Sub : End If
        If Val(id_EMPLEADO) = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Situacion_Personal_Imprimir.ShowDialog()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call VERIFICAR_ACCESOS("186") : If HAY_ACCESO = False Then : Exit Sub : End If
        If Val(msID_M_A) = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        If msID_SIT_P = 38 Or msID_SIT_P = 39 Then
            MsgBox("Las Situaciones [AJUSTE MENSUAL (+)] y [AJUSTE MENSUAL (-)] no pueden ser modificadas", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Situacion_Personal_Editar.ShowDialog()
        If CERRAR = False Then
            FECHA01 = Me.DateTimePicker1.Text
            FECHA1 = Mid(FECHA01, 1, 10)
            FECHA02 = Me.DateTimePicker2.Text
            FECHA2 = Mid(FECHA02, 1, 10)
            If isuID_TIPO_USUARIO = 5 Then
                CADENAsql = "SET LANGUAGE ESPAÑOL Select ID, NIVEL1, NIVEL2, NIVEL3, NIVEL4, NIVEL5, ORD, ID_M_A, ID_M_P, CODIGO, UPPER(DATENAME (dw, DIA)) AS DIALETRAS, DIA, ID_T_SIT_P, D_T_SIT_P, ID_SIT_P, D_SIT_P, SIGLA, VALOR_SIT, VALOR_DIA, OBSERVACIONES, DETALLE_1, DETALLE_2, DETALLE_3, D_ADMIN, USUARIO_ACT, FECHA_ACT from [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE CODIGO = '" & Me.TextBox2.Text & "' AND DIA BETWEEN '" & FECHA1 & "' AND '" & FECHA2 & "' AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY CODIGO, DIA SET LANGUAGE ENGLISH"
            Else
                CADENAsql = "SET LANGUAGE ESPAÑOL Select ID, NIVEL1, NIVEL2, NIVEL3, NIVEL4, NIVEL5, ORD, ID_M_A, ID_M_P, CODIGO, UPPER(DATENAME (dw, DIA)) AS DIALETRAS, DIA, ID_T_SIT_P, D_T_SIT_P, ID_SIT_P, D_SIT_P, SIGLA, VALOR_SIT, VALOR_DIA, OBSERVACIONES, DETALLE_1, DETALLE_2, DETALLE_3, D_ADMIN, USUARIO_ACT, FECHA_ACT from [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE CODIGO = '" & Me.TextBox2.Text & "' AND DIA BETWEEN '" & FECHA1 & "' AND '" & FECHA2 & "' ORDER BY CODIGO, DIA SET LANGUAGE ENGLISH"
            End If

            RemoveHandler Me.DGV.SelectionChanged, AddressOf DGV_SelectionChanged
            RemoveHandler Me.DGV.RowPrePaint, AddressOf DGV_RowPrePaint
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
            '--------------------------
            Call DIAS_SUBSIDIOS()
            Call DIAS_AI()
            Call DIAS_V()
            Call DIAS_OM()
            Label11.Text = vOM + vAI
            '--------------------------
            AddHandler Me.DGV.RowPrePaint, AddressOf DGV_RowPrePaint
            AddHandler Me.DGV.SelectionChanged, AddressOf DGV_SelectionChanged
            Call CARGAR_DATAGRIDVIEW_15()
            Call ARMAR_DATAGRIDVIEW_15()
        End If
    End Sub
    Private Sub DGV_Click(sender As Object, e As EventArgs) Handles DGV.Click
        If DGV.SelectedCells.Count <> 0 Then
            Me.Label2.Text = DGV.RowCount & " Registro(s) Existente(s) ---> Seleccionados: " & DGV.SelectedRows.Count
            Me.Label16.Text = "Última Actualización del registro (Fecha y Usuario) ---> " & Mid(msFECHA_ACT, 1, 10) & ", " & msUSUARIO_ACT
            Dim FILA = DGV.CurrentRow.Index
            If Not IsDBNull(Me.DGV.Item(0, FILA).Value) Then : msID = Me.DGV.Item(0, FILA).Value : Else : msID = 0 : End If
            If Not IsDBNull(Me.DGV.Item(1, FILA).Value) Then : msNIVEL1 = Me.DGV.Item(1, FILA).Value : Else : msNIVEL1 = "" : End If
            If Not IsDBNull(Me.DGV.Item(2, FILA).Value) Then : msNIVEL2 = Me.DGV.Item(2, FILA).Value : Else : msNIVEL2 = "" : End If
            If Not IsDBNull(Me.DGV.Item(3, FILA).Value) Then : msNIVEL3 = Me.DGV.Item(3, FILA).Value : Else : msNIVEL3 = "" : End If
            If Not IsDBNull(Me.DGV.Item(4, FILA).Value) Then : msNIVEL4 = Me.DGV.Item(4, FILA).Value : Else : msNIVEL4 = "" : End If
            If Not IsDBNull(Me.DGV.Item(5, FILA).Value) Then : msNIVEL5 = Me.DGV.Item(5, FILA).Value : Else : msNIVEL5 = "" : End If
            If Not IsDBNull(Me.DGV.Item(6, FILA).Value) Then : msORD = Me.DGV.Item(6, FILA).Value : Else : msORD = 0 : End If
            If Not IsDBNull(Me.DGV.Item(7, FILA).Value) Then : msID_M_A = Me.DGV.Item(7, FILA).Value : Else : msID_M_A = 0 : End If
            If Not IsDBNull(Me.DGV.Item(8, FILA).Value) Then : msID_M_P = Me.DGV.Item(8, FILA).Value : Else : msID_M_P = 0 : End If
            If Not IsDBNull(Me.DGV.Item(9, FILA).Value) Then : msCODIGO = Me.DGV.Item(9, FILA).Value : Else : msCODIGO = "" : End If
            If Not IsDBNull(Me.DGV.Item(10, FILA).Value) Then : msDIAL = Me.DGV.Item(10, FILA).Value : Else : msDIAL = "" : End If
            If Not IsDBNull(Me.DGV.Item(11, FILA).Value) Then : msDIA = Me.DGV.Item(11, FILA).Value : Else : msDIA = "" : End If
            If Not IsDBNull(Me.DGV.Item(12, FILA).Value) Then : msID_T_SIT_P = Me.DGV.Item(12, FILA).Value : Else : msID_T_SIT_P = 0 : End If
            If Not IsDBNull(Me.DGV.Item(13, FILA).Value) Then : msD_T_SIT_P = Me.DGV.Item(13, FILA).Value : Else : msD_T_SIT_P = "" : End If
            If Not IsDBNull(Me.DGV.Item(14, FILA).Value) Then : msID_SIT_P = Me.DGV.Item(14, FILA).Value : Else : msID_SIT_P = 0 : End If
            If Not IsDBNull(Me.DGV.Item(15, FILA).Value) Then : msD_SIT_P = Me.DGV.Item(15, FILA).Value : Else : msD_SIT_P = "" : End If
            If Not IsDBNull(Me.DGV.Item(16, FILA).Value) Then : msSIGLA = Me.DGV.Item(16, FILA).Value : Else : msSIGLA = "" : End If
            If Not IsDBNull(Me.DGV.Item(17, FILA).Value) Then : msVALOR_SIT = Me.DGV.Item(17, FILA).Value : Else : msVALOR_SIT = "" : End If
            If Not IsDBNull(Me.DGV.Item(18, FILA).Value) Then : msVALOR_DIA = Me.DGV.Item(18, FILA).Value : Else : msVALOR_DIA = "" : End If
            If Not IsDBNull(Me.DGV.Item(19, FILA).Value) Then : msOBSERVACIONES = Me.DGV.Item(19, FILA).Value : Else : msOBSERVACIONES = "" : End If
            If Not IsDBNull(Me.DGV.Item(20, FILA).Value) Then : msDETALLE_1 = Me.DGV.Item(20, FILA).Value : Else : msDETALLE_1 = "" : End If
            If Not IsDBNull(Me.DGV.Item(21, FILA).Value) Then : msDETALLE_2 = Me.DGV.Item(21, FILA).Value : Else : msDETALLE_2 = "" : End If
            If Not IsDBNull(Me.DGV.Item(22, FILA).Value) Then : msDETALLE_3 = Me.DGV.Item(22, FILA).Value : Else : msDETALLE_3 = "" : End If
            If Not IsDBNull(Me.DGV.Item(23, FILA).Value) Then : msD_ADMIN = Me.DGV.Item(23, FILA).Value : Else : msD_ADMIN = "" : End If
            If Not IsDBNull(Me.DGV.Item(24, FILA).Value) Then : msUSUARIO_ACT = Me.DGV.Item(24, FILA).Value : Else : msUSUARIO_ACT = "" : End If
            If Not IsDBNull(Me.DGV.Item(25, FILA).Value) Then : msFECHA_ACT = Me.DGV.Item(25, FILA).Value : Else : msFECHA_ACT = "" : End If

            Call CARGAR_DATAGRIDVIEW_15()
            Call ARMAR_DATAGRIDVIEW_15()
        Else
            Me.Label16.Text = "Última Actualización del registro (Fecha y Usuario) ---> Actualmente sin Actualizacion"
        End If
    End Sub
    Private Sub DGV_SelectionChanged(sender As Object, e As EventArgs) Handles DGV.SelectionChanged
        If DGV.SelectedCells.Count <> 0 Then
            Me.Label2.Text = DGV.RowCount & " Registro(s) Existente(s) ---> Seleccionados: " & DGV.SelectedRows.Count
            Me.Label16.Text = "Última Actualización del registro (Fecha y Usuario) ---> " & Mid(msFECHA_ACT, 1, 10) & ", " & msUSUARIO_ACT
            Dim FILA = DGV.CurrentRow.Index
            If Not IsDBNull(Me.DGV.Item(0, FILA).Value) Then : msID = Me.DGV.Item(0, FILA).Value : Else : msID = 0 : End If
            If Not IsDBNull(Me.DGV.Item(1, FILA).Value) Then : msNIVEL1 = Me.DGV.Item(1, FILA).Value : Else : msNIVEL1 = "" : End If
            If Not IsDBNull(Me.DGV.Item(2, FILA).Value) Then : msNIVEL2 = Me.DGV.Item(2, FILA).Value : Else : msNIVEL2 = "" : End If
            If Not IsDBNull(Me.DGV.Item(3, FILA).Value) Then : msNIVEL3 = Me.DGV.Item(3, FILA).Value : Else : msNIVEL3 = "" : End If
            If Not IsDBNull(Me.DGV.Item(4, FILA).Value) Then : msNIVEL4 = Me.DGV.Item(4, FILA).Value : Else : msNIVEL4 = "" : End If
            If Not IsDBNull(Me.DGV.Item(5, FILA).Value) Then : msNIVEL5 = Me.DGV.Item(5, FILA).Value : Else : msNIVEL5 = "" : End If
            If Not IsDBNull(Me.DGV.Item(6, FILA).Value) Then : msORD = Me.DGV.Item(6, FILA).Value : Else : msORD = 0 : End If
            If Not IsDBNull(Me.DGV.Item(7, FILA).Value) Then : msID_M_A = Me.DGV.Item(7, FILA).Value : Else : msID_M_A = 0 : End If
            If Not IsDBNull(Me.DGV.Item(8, FILA).Value) Then : msID_M_P = Me.DGV.Item(8, FILA).Value : Else : msID_M_P = 0 : End If
            If Not IsDBNull(Me.DGV.Item(9, FILA).Value) Then : msCODIGO = Me.DGV.Item(9, FILA).Value : Else : msCODIGO = "" : End If
            If Not IsDBNull(Me.DGV.Item(10, FILA).Value) Then : msDIAL = Me.DGV.Item(10, FILA).Value : Else : msDIAL = "" : End If
            If Not IsDBNull(Me.DGV.Item(11, FILA).Value) Then : msDIA = Me.DGV.Item(11, FILA).Value : Else : msDIA = "" : End If
            If Not IsDBNull(Me.DGV.Item(12, FILA).Value) Then : msID_T_SIT_P = Me.DGV.Item(12, FILA).Value : Else : msID_T_SIT_P = 0 : End If
            If Not IsDBNull(Me.DGV.Item(13, FILA).Value) Then : msD_T_SIT_P = Me.DGV.Item(13, FILA).Value : Else : msD_T_SIT_P = "" : End If
            If Not IsDBNull(Me.DGV.Item(14, FILA).Value) Then : msID_SIT_P = Me.DGV.Item(14, FILA).Value : Else : msID_SIT_P = 0 : End If
            If Not IsDBNull(Me.DGV.Item(15, FILA).Value) Then : msD_SIT_P = Me.DGV.Item(15, FILA).Value : Else : msD_SIT_P = "" : End If
            If Not IsDBNull(Me.DGV.Item(16, FILA).Value) Then : msSIGLA = Me.DGV.Item(16, FILA).Value : Else : msSIGLA = "" : End If
            If Not IsDBNull(Me.DGV.Item(17, FILA).Value) Then : msVALOR_SIT = Me.DGV.Item(17, FILA).Value : Else : msVALOR_SIT = "" : End If
            If Not IsDBNull(Me.DGV.Item(18, FILA).Value) Then : msVALOR_DIA = Me.DGV.Item(18, FILA).Value : Else : msVALOR_DIA = "" : End If
            If Not IsDBNull(Me.DGV.Item(19, FILA).Value) Then : msOBSERVACIONES = Me.DGV.Item(19, FILA).Value : Else : msOBSERVACIONES = "" : End If
            If Not IsDBNull(Me.DGV.Item(20, FILA).Value) Then : msDETALLE_1 = Me.DGV.Item(20, FILA).Value : Else : msDETALLE_1 = "" : End If
            If Not IsDBNull(Me.DGV.Item(21, FILA).Value) Then : msDETALLE_2 = Me.DGV.Item(21, FILA).Value : Else : msDETALLE_2 = "" : End If
            If Not IsDBNull(Me.DGV.Item(22, FILA).Value) Then : msDETALLE_3 = Me.DGV.Item(22, FILA).Value : Else : msDETALLE_3 = "" : End If
            If Not IsDBNull(Me.DGV.Item(23, FILA).Value) Then : msD_ADMIN = Me.DGV.Item(23, FILA).Value : Else : msD_ADMIN = "" : End If
            If Not IsDBNull(Me.DGV.Item(24, FILA).Value) Then : msUSUARIO_ACT = Me.DGV.Item(24, FILA).Value : Else : msUSUARIO_ACT = "" : End If
            If Not IsDBNull(Me.DGV.Item(25, FILA).Value) Then : msFECHA_ACT = Me.DGV.Item(25, FILA).Value : Else : msFECHA_ACT = "" : End If
            Call CARGAR_DATAGRIDVIEW_15()
            Call ARMAR_DATAGRIDVIEW_15()
        Else
            Me.Label16.Text = "Última Actualización del registro (Fecha y Usuario) ---> Actualmente sin Actualizacion"
        End If
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        Dim I As Integer
        For I = 0 To Me.DGV.RowCount - 1
            If Me.DGV.Item(7, I).Value = msID_M_A Then
                Me.DGV.Rows(I).Selected = True
                Me.DGV.CurrentCell = Me.DGV.Item(7, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub DGV_DoubleClick(sender As Object, e As EventArgs) Handles DGV.DoubleClick
        Call VERIFICAR_ACCESOS("186") : If HAY_ACCESO = False Then : Exit Sub : End If
        If Val(msID_M_A) = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        If msID_SIT_P = 38 Or msID_SIT_P = 39 Then
            MsgBox("Las Situaciones [AJUSTE MENSUAL (+)] y [AJUSTE MENSUAL (-)] no pueden ser modificadas", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Situacion_Personal_Editar.ShowDialog()

        If CERRAR = False Then
            FECHA01 = Me.DateTimePicker1.Text
            FECHA1 = Mid(FECHA01, 1, 10)
            FECHA02 = Me.DateTimePicker2.Text
            FECHA2 = Mid(FECHA02, 1, 10)
            If isuID_TIPO_USUARIO = 5 Then
                CADENAsql = "SET LANGUAGE ESPAÑOL Select ID, NIVEL1, NIVEL2, NIVEL3, NIVEL4, NIVEL5, ORD, ID_M_A, ID_M_P, CODIGO, UPPER(DATENAME (dw, DIA)) AS DIALETRAS, DIA, ID_T_SIT_P, D_T_SIT_P, ID_SIT_P, D_SIT_P, SIGLA, VALOR_SIT, VALOR_DIA, OBSERVACIONES, DETALLE_1, DETALLE_2, DETALLE_3, D_ADMIN, USUARIO_ACT, FECHA_ACT from [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE CODIGO = '" & Me.TextBox2.Text & "' AND DIA BETWEEN '" & FECHA1 & "' AND '" & FECHA2 & "' AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY CODIGO, DIA SET LANGUAGE ENGLISH"
            Else
                CADENAsql = "SET LANGUAGE ESPAÑOL Select ID, NIVEL1, NIVEL2, NIVEL3, NIVEL4, NIVEL5, ORD, ID_M_A, ID_M_P, CODIGO, UPPER(DATENAME (dw, DIA)) AS DIALETRAS, DIA, ID_T_SIT_P, D_T_SIT_P, ID_SIT_P, D_SIT_P, SIGLA, VALOR_SIT, VALOR_DIA, OBSERVACIONES, DETALLE_1, DETALLE_2, DETALLE_3, D_ADMIN, USUARIO_ACT, FECHA_ACT from [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE CODIGO = '" & Me.TextBox2.Text & "' AND DIA BETWEEN '" & FECHA1 & "' AND '" & FECHA2 & "' ORDER BY CODIGO, DIA SET LANGUAGE ENGLISH"
            End If
            RemoveHandler Me.DGV.SelectionChanged, AddressOf DGV_SelectionChanged
            RemoveHandler Me.DGV.RowPrePaint, AddressOf DGV_RowPrePaint
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
            '--------------------------
            Call DIAS_SUBSIDIOS()
            Call DIAS_AI()
            Call DIAS_V()
            Call DIAS_OM()
            Label11.Text = vOM + vAI
            '--------------------------
            AddHandler Me.DGV.RowPrePaint, AddressOf DGV_RowPrePaint
            AddHandler Me.DGV.SelectionChanged, AddressOf DGV_SelectionChanged
            Call CARGAR_DATAGRIDVIEW_15()
            Call ARMAR_DATAGRIDVIEW_15()
        End If
    End Sub
    Private Sub DGV_RowPrePaint(sender As Object, e As DataGridViewRowPrePaintEventArgs) Handles DGV.RowPrePaint
        Dim style As New DataGridViewCellStyle
        style.Font = New Font(DGV.Font, FontStyle.Bold)
        For Each Row As DataGridViewRow In DGV.Rows
            If Row.Cells(15).Value = "AUSENCIA INJUSTIFICADA" Or Row.Cells(15).Value = "OMISION DE MARCA" Then
                Row.DefaultCellStyle.ForeColor = Color.Red
            End If
            If Row.Cells(15).Value = "VACACIONES" Then
                Row.DefaultCellStyle.ForeColor = Color.Blue
            End If
            If Row.Cells(15).Value = "SUBSIDIO" Then
                Row.DefaultCellStyle.ForeColor = Color.DarkOrange
            End If
            If Row.Cells(15).Value = "AJUSTE MENSUAL (+)" Then
                Row.DefaultCellStyle.ForeColor = Color.Red
            End If
            If Row.Cells(15).Value = "AJUSTE MENSUAL (-)" Then
                Row.DefaultCellStyle.ForeColor = Color.red
            End If
        Next
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If Me.CheckBox1.Checked = True Then
            Me.DGV.MultiSelect = True
        Else
            Me.DGV.MultiSelect = False
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call VERIFICAR_ACCESOS("188") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Situacion_Personal_Ajustes.ShowDialog()
    End Sub
    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        Me.DateTimePicker20.Value = Me.DateTimePicker2.Value
        FECHA02 = Me.DateTimePicker2.Text
        FECHA2 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA02 + "', 105), 23)"
        If Len(Me.TextBox2.Text) = 0 Or (Me.TextBox2.Text = "0000000000") Then
            Exit Sub
        End If
        FECHA2 = Mid(FECHA02, 1, 10)
        If isuID_TIPO_USUARIO = 5 Then
            CADENAsql = "SET LANGUAGE ESPAÑOL Select ID, NIVEL1, NIVEL2, NIVEL3, NIVEL4, NIVEL5, ORD, ID_M_A, ID_M_P, CODIGO, UPPER(DATENAME (dw, DIA)) AS DIALETRAS, DIA, ID_T_SIT_P, D_T_SIT_P, ID_SIT_P, D_SIT_P, SIGLA, VALOR_SIT, VALOR_DIA, OBSERVACIONES, DETALLE_1, DETALLE_2, DETALLE_3, D_ADMIN, USUARIO_ACT, FECHA_ACT from [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE CODIGO = '" & Me.TextBox2.Text & "' AND DIA BETWEEN '" & FECHA1 & "' AND '" & FECHA2 & "' AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY CODIGO, DIA SET LANGUAGE ENGLISH"
        Else
            CADENAsql = "SET LANGUAGE ESPAÑOL Select ID, NIVEL1, NIVEL2, NIVEL3, NIVEL4, NIVEL5, ORD, ID_M_A, ID_M_P, CODIGO, UPPER(DATENAME (dw, DIA)) AS DIALETRAS, DIA, ID_T_SIT_P, D_T_SIT_P, ID_SIT_P, D_SIT_P, SIGLA, VALOR_SIT, VALOR_DIA, OBSERVACIONES, DETALLE_1, DETALLE_2, DETALLE_3, D_ADMIN, USUARIO_ACT, FECHA_ACT from [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE CODIGO = '" & Me.TextBox2.Text & "' AND DIA BETWEEN '" & FECHA1 & "' AND '" & FECHA2 & "' ORDER BY CODIGO, DIA SET LANGUAGE ENGLISH"
        End If
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
        '--------------------------
        Call DIAS_SUBSIDIOS()
        Call DIAS_AI()
        Call DIAS_V()
        Call DIAS_OM()
        Label11.Text = vOM + vAI
        '--------------------------
        Call CARGAR_DATAGRIDVIEW_15()
        Call ARMAR_DATAGRIDVIEW_15()
    End Sub
    Private Sub TextBox2_DoubleClick(sender As Object, e As EventArgs) Handles TextBox2.DoubleClick
        Call VERIFICAR_ACCESOS("184") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Situacion_Personal_Buscar.ShowDialog()
        If CE <> "" Then
            Me.TextBox2.Text = CE

            CADENAsql = "SELECT ID_M_P, CODIGO, APELLIDOS, NOMBRES, N_CEDULA, FEC_ING_PAME, ID_ESTADO_P, ID_N_PROFESIONAL, D_T_SEXO FROM [dbo].[VISTA MAESTRO DE PERSONAS] WHERE CODIGO = '" & Me.TextBox2.Text & "' AND ID_ESTADO_P = 1"
            Call BUSCAR_DATO_1()
            If DATO_1 = True Then
                FECHA01 = Mid(Me.DateTimePicker1.Text, 1, 10)
                FECHA1 = Mid(FECHA01, 1, 10)
                FECHA02 = Mid(Me.DateTimePicker2.Text, 1, 10)
                FECHA2 = Mid(FECHA02, 1, 10)
                If isuID_TIPO_USUARIO = 5 Then
                    CADENAsql = "SET LANGUAGE ESPAÑOL Select ID, NIVEL1, NIVEL2, NIVEL3, NIVEL4, NIVEL5, ORD, ID_M_A, ID_M_P, CODIGO, UPPER(DATENAME (dw, DIA)) AS DIALETRAS, DIA, ID_T_SIT_P, D_T_SIT_P, ID_SIT_P, D_SIT_P, SIGLA, VALOR_SIT, VALOR_DIA, OBSERVACIONES, DETALLE_1, DETALLE_2, DETALLE_3, D_ADMIN, USUARIO_ACT, FECHA_ACT from [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE CODIGO = '" & Me.TextBox2.Text & "' AND DIA BETWEEN '" & FECHA1 & "' AND '" & FECHA2 & "' AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY CODIGO, DIA SET LANGUAGE ENGLISH"
                Else
                    CADENAsql = "SET LANGUAGE ESPAÑOL Select ID, NIVEL1, NIVEL2, NIVEL3, NIVEL4, NIVEL5, ORD, ID_M_A, ID_M_P, CODIGO, UPPER(DATENAME (dw, DIA)) AS DIALETRAS, DIA, ID_T_SIT_P, D_T_SIT_P, ID_SIT_P, D_SIT_P, SIGLA, VALOR_SIT, VALOR_DIA, OBSERVACIONES, DETALLE_1, DETALLE_2, DETALLE_3, D_ADMIN, USUARIO_ACT, FECHA_ACT from [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE CODIGO = '" & Me.TextBox2.Text & "' AND DIA BETWEEN '" & FECHA1 & "' AND '" & FECHA2 & "' ORDER BY CODIGO, DIA SET LANGUAGE ENGLISH"
                End If
                Call CARGAR_DATAGRIDVIEW()
                Call ARMAR_DATAGRIDVIEW()
                '--------------------------
                Call DIAS_SUBSIDIOS()
                Call DIAS_AI()
                Call DIAS_V()
                Call DIAS_OM()
                Label11.Text = vOM + vAI
                '--------------------------
                Call BUSCAR_CARGO()
                Call CARGAR_DATAGRIDVIEW_15()
                Call ARMAR_DATAGRIDVIEW_15()
            Else
                MsgBox("El código que se digitó no es válido", vbInformation, "Mensaje del Sistema")
                Me.TextBox2.SelectAll()
                Me.TextBox2.Focus()
                Exit Sub
            End If
            Me.TextBox2.SelectAll()
            Me.TextBox2.Focus()

        Else
            Me.TextBox2.Text = "0000000000"
            Me.TextBox2.SelectAll()
            Me.TextBox2.Focus()
        End If
    End Sub
    Dim vOM As Integer
    Private Sub DIAS_OM()
        Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker1.Value, 1, 10) + "', 105), 23)"
        Dim F2 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker2.Value, 1, 10) + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable3 As New DataTable
        Dim MiDataAdapter3 As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter3 = New SqlDataAdapter("SELECT SUM(VALOR_SIT) FROM [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE ([CODIGO] = '" & Me.TextBox2.Text & "') AND (ID_SIT_P = 20) AND (DIA BETWEEN " & F1 & " AND " & F2 & ")", Conexion.CxRRHH)
            MiDataTable3.Clear()
            MiDataAdapter3.Fill(MiDataTable3)
            If MiDataTable3.Rows.Count <> 0 Then
                vOM = Math.Abs(Val(MiDataTable3.Rows(0).Item(0).ToString))
            Else
                vOM = "0"
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub DIAS_V()
        Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker1.Value, 1, 10) + "', 105), 23)"
        Dim F2 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker2.Value, 1, 10) + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable3 As New DataTable
        Dim MiDataAdapter3 As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter3 = New SqlDataAdapter("SELECT SUM(VALOR_SIT) FROM [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE ([CODIGO] = '" & Me.TextBox2.Text & "') AND (ID_SIT_P = 17) AND (DIA BETWEEN " & F1 & " AND " & F2 & ")", Conexion.CxRRHH)
            MiDataTable3.Clear()
            MiDataAdapter3.Fill(MiDataTable3)
            If MiDataTable3.Rows.Count <> 0 Then
                Me.Label13.Text = Math.Abs(Val(MiDataTable3.Rows(0).Item(0).ToString))
            Else
                Me.Label13.Text = "0"
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim vAI As Integer
    Private Sub DIAS_AI()
        Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker1.Value, 1, 10) + "', 105), 23)"
        Dim F2 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker2.Value, 1, 10) + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable3 As New DataTable
        Dim MiDataAdapter3 As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter3 = New SqlDataAdapter("SELECT SUM(VALOR_SIT) FROM [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE ([CODIGO] = '" & Me.TextBox2.Text & "') AND (ID_SIT_P = 3) AND (DIA BETWEEN " & F1 & " AND " & F2 & ")", Conexion.CxRRHH)
            MiDataTable3.Clear()
            MiDataAdapter3.Fill(MiDataTable3)
            If MiDataTable3.Rows.Count <> 0 Then
                vAI = Math.Abs(Val(MiDataTable3.Rows(0).Item(0).ToString))
            Else
                vAI = "0"
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call VERIFICAR_ACCESOS("185") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Situacion_Personal_Proy.ShowDialog()
        FECHA01 = Me.DateTimePicker1.Text
        FECHA1 = Mid(FECHA01, 1, 10)
        FECHA02 = Me.DateTimePicker2.Text
        FECHA2 = Mid(FECHA02, 1, 10)
        If isuID_TIPO_USUARIO = 5 Then
            CADENAsql = "SET LANGUAGE ESPAÑOL Select ID, NIVEL1, NIVEL2, NIVEL3, NIVEL4, NIVEL5, ORD, ID_M_A, ID_M_P, CODIGO, UPPER(DATENAME (dw, DIA)) AS DIALETRAS, DIA, ID_T_SIT_P, D_T_SIT_P, ID_SIT_P, D_SIT_P, SIGLA, VALOR_SIT, VALOR_DIA, OBSERVACIONES, DETALLE_1, DETALLE_2, DETALLE_3, D_ADMIN, USUARIO_ACT, FECHA_ACT from [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE CODIGO = '" & Me.TextBox2.Text & "' AND DIA BETWEEN '" & FECHA1 & "' AND '" & FECHA2 & "' AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY CODIGO, DIA SET LANGUAGE ENGLISH"
        Else
            CADENAsql = "SET LANGUAGE ESPAÑOL Select ID, NIVEL1, NIVEL2, NIVEL3, NIVEL4, NIVEL5, ORD, ID_M_A, ID_M_P, CODIGO, UPPER(DATENAME (dw, DIA)) AS DIALETRAS, DIA, ID_T_SIT_P, D_T_SIT_P, ID_SIT_P, D_SIT_P, SIGLA, VALOR_SIT, VALOR_DIA, OBSERVACIONES, DETALLE_1, DETALLE_2, DETALLE_3, D_ADMIN, USUARIO_ACT, FECHA_ACT from [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE CODIGO = '" & Me.TextBox2.Text & "' AND DIA BETWEEN '" & FECHA1 & "' AND '" & FECHA2 & "' ORDER BY CODIGO, DIA SET LANGUAGE ENGLISH"
        End If

        RemoveHandler Me.DGV.SelectionChanged, AddressOf DGV_SelectionChanged
        RemoveHandler Me.DGV.RowPrePaint, AddressOf DGV_RowPrePaint
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
        Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        '--------------------------
        Call DIAS_SUBSIDIOS()
        Call DIAS_AI()
        Call DIAS_V()
        Call DIAS_OM()
        Label11.Text = vOM + vAI
        '--------------------------
        AddHandler Me.DGV.RowPrePaint, AddressOf DGV_RowPrePaint
        AddHandler Me.DGV.SelectionChanged, AddressOf DGV_SelectionChanged
        Call CARGAR_DATAGRIDVIEW_15()
        Call ARMAR_DATAGRIDVIEW_15()
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Call VERIFICAR_ACCESOS("310") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Situacion_Personal_Agregar.ShowDialog()
        If CERRAR = False Then
            FECHA01 = Me.DateTimePicker1.Text
            FECHA1 = Mid(FECHA01, 1, 10)
            FECHA02 = Me.DateTimePicker2.Text
            FECHA2 = Mid(FECHA02, 1, 10)
            If isuID_TIPO_USUARIO = 5 Then
                CADENAsql = "SET LANGUAGE ESPAÑOL Select ID, NIVEL1, NIVEL2, NIVEL3, NIVEL4, NIVEL5, ORD, ID_M_A, ID_M_P, CODIGO, UPPER(DATENAME (dw, DIA)) AS DIALETRAS, DIA, ID_T_SIT_P, D_T_SIT_P, ID_SIT_P, D_SIT_P, SIGLA, VALOR_SIT, VALOR_DIA, OBSERVACIONES, DETALLE_1, DETALLE_2, DETALLE_3, D_ADMIN, USUARIO_ACT, FECHA_ACT from [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE CODIGO = '" & Me.TextBox2.Text & "' AND DIA BETWEEN '" & FECHA1 & "' AND '" & FECHA2 & "' AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY CODIGO, DIA SET LANGUAGE ENGLISH"
            Else
                CADENAsql = "SET LANGUAGE ESPAÑOL Select ID, NIVEL1, NIVEL2, NIVEL3, NIVEL4, NIVEL5, ORD, ID_M_A, ID_M_P, CODIGO, UPPER(DATENAME (dw, DIA)) AS DIALETRAS, DIA, ID_T_SIT_P, D_T_SIT_P, ID_SIT_P, D_SIT_P, SIGLA, VALOR_SIT, VALOR_DIA, OBSERVACIONES, DETALLE_1, DETALLE_2, DETALLE_3, D_ADMIN, USUARIO_ACT, FECHA_ACT from [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE CODIGO = '" & Me.TextBox2.Text & "' AND DIA BETWEEN '" & FECHA1 & "' AND '" & FECHA2 & "' ORDER BY CODIGO, DIA SET LANGUAGE ENGLISH"
            End If

            RemoveHandler Me.DGV.SelectionChanged, AddressOf DGV_SelectionChanged
            RemoveHandler Me.DGV.RowPrePaint, AddressOf DGV_RowPrePaint
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
            '--------------------------
            Call DIAS_SUBSIDIOS()
            Call DIAS_AI()
            Call DIAS_V()
            Call DIAS_OM()
            Label11.Text = vOM + vAI
            '--------------------------
            AddHandler Me.DGV.RowPrePaint, AddressOf DGV_RowPrePaint
            AddHandler Me.DGV.SelectionChanged, AddressOf DGV_SelectionChanged
            Call CARGAR_DATAGRIDVIEW_15()
            Call ARMAR_DATAGRIDVIEW_15()
        End If
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Call VERIFICAR_ACCESOS("311") : If HAY_ACCESO = False Then : Exit Sub : End If
        If Val(msID_M_A) = 0 Then
            MsgBox("Debe seleccionar el registro que desea Eliminar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Esta seguro de Eliminar el registro seleccionado?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            Call ELIMINAR_REGISTRO()
            If CERRAR = False Then
                FECHA01 = Me.DateTimePicker1.Text
                FECHA1 = Mid(FECHA01, 1, 10)
                FECHA02 = Me.DateTimePicker2.Text
                FECHA2 = Mid(FECHA02, 1, 10)
                If isuID_TIPO_USUARIO = 5 Then
                    CADENAsql = "SET LANGUAGE ESPAÑOL Select ID, NIVEL1, NIVEL2, NIVEL3, NIVEL4, NIVEL5, ORD, ID_M_A, ID_M_P, CODIGO, UPPER(DATENAME (dw, DIA)) AS DIALETRAS, DIA, ID_T_SIT_P, D_T_SIT_P, ID_SIT_P, D_SIT_P, SIGLA, VALOR_SIT, VALOR_DIA, OBSERVACIONES, DETALLE_1, DETALLE_2, DETALLE_3, D_ADMIN, USUARIO_ACT, FECHA_ACT from [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE CODIGO = '" & Me.TextBox2.Text & "' AND DIA BETWEEN '" & FECHA1 & "' AND '" & FECHA2 & "' AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY CODIGO, DIA SET LANGUAGE ENGLISH"
                Else
                    CADENAsql = "SET LANGUAGE ESPAÑOL Select ID, NIVEL1, NIVEL2, NIVEL3, NIVEL4, NIVEL5, ORD, ID_M_A, ID_M_P, CODIGO, UPPER(DATENAME (dw, DIA)) AS DIALETRAS, DIA, ID_T_SIT_P, D_T_SIT_P, ID_SIT_P, D_SIT_P, SIGLA, VALOR_SIT, VALOR_DIA, OBSERVACIONES, DETALLE_1, DETALLE_2, DETALLE_3, D_ADMIN, USUARIO_ACT, FECHA_ACT from [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE CODIGO = '" & Me.TextBox2.Text & "' AND DIA BETWEEN '" & FECHA1 & "' AND '" & FECHA2 & "' ORDER BY CODIGO, DIA SET LANGUAGE ENGLISH"
                End If

                RemoveHandler Me.DGV.SelectionChanged, AddressOf DGV_SelectionChanged
                RemoveHandler Me.DGV.RowPrePaint, AddressOf DGV_RowPrePaint
                Call CARGAR_DATAGRIDVIEW()
                Call ARMAR_DATAGRIDVIEW()
                Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
                '--------------------------
                Call DIAS_SUBSIDIOS()
                Call DIAS_AI()
                Call DIAS_V()
                Call DIAS_OM()
                Label11.Text = vOM + vAI
                '--------------------------
                AddHandler Me.DGV.RowPrePaint, AddressOf DGV_RowPrePaint
                AddHandler Me.DGV.SelectionChanged, AddressOf DGV_SelectionChanged
                Call CARGAR_DATAGRIDVIEW_15()
                Call ARMAR_DATAGRIDVIEW_15()
            End If
        End If
    End Sub
    Private Sub ELIMINAR_REGISTRO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE SITUACION DE PERSONAL] WHERE ID_M_A = " & CInt(msID_M_A) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub DIAS_SUBSIDIOS()
        Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker1.Value, 1, 10) + "', 105), 23)"
        Dim F2 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker2.Value, 1, 10) + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable3 As New DataTable
        Dim MiDataAdapter3 As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter3 = New SqlDataAdapter("SELECT COUNT(VALOR_SIT) FROM [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE ([CODIGO] = '" & Me.TextBox2.Text & "') AND (ID_SIT_P = 14) AND (DIA BETWEEN " & F1 & " AND " & F2 & ")", Conexion.CxRRHH)
            MiDataTable3.Clear()
            MiDataAdapter3.Fill(MiDataTable3)
            If MiDataTable3.Rows.Count <> 0 Then
                Me.Label17.Text = Math.Abs(Val(MiDataTable3.Rows(0).Item(0).ToString))
            Else
                Me.Label17.Text = "0"
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
End Class