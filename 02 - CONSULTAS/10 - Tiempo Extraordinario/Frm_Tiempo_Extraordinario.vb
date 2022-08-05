Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Public Class Frm_Tiempo_Extraordinario
    Private Sub Frm_Tiempo_Extraordinario_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            TIEMPO_EXTRAORDINARIO = False
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Tiempo_Extraordinario_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Call Clases.BUSCAR_FECHA_Y_HORA()
        Me.DateTimePicker1.Value = Mid(FECHA_SERVIDOR, 1, 10)
        Me.DateTimePicker2.Value = Mid(FECHA_SERVIDOR, 1, 10)
        TIEMPO_EXTRAORDINARIO = True
        Me.DGV1.DataSource = Nothing
        Me.TextBox1.Focus()
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        TIEMPO_EXTRAORDINARIO = False
        Me.Close()
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
                CADENAsql = "SELECT ID_M_P, CODIGO, APELLIDOS, NOMBRES, N_CEDULA, FEC_ING_PAME, ID_ESTADO_P, ID_N_PROFESIONAL, ID_T_SEXO, PREFIJO FROM [dbo].[VISTA MAESTRO DE PERSONAS] WHERE CODIGO = '" & Me.TextBox1.Text & "' AND ID_ESTADO_P = 1"
                Call BUSCAR_DATO_1()
                If DATO_1 = True Then
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
                Me.TextBox2.Text = MiDataTable.Rows(0).Item(2).ToString & " " & MiDataTable.Rows(0).Item(3).ToString

                DATO_1 = True
            Else
                Me.TextBox2.Text = ""
                NP = 0
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.TextBox1.Text <> "" Then
            If Mid(Me.DateTimePicker2.Value, 1, 10) < Mid(Me.DateTimePicker1.Value, 1, 10) Then
                MsgBox("La Fecha Final no puede ser menor a la Fecha Inicial", vbInformation, "Mensaje del Sistema")
                Me.DateTimePicker1.Focus()
            Else
                Call BORRAR_DATOS_USUARIO()
                Call EJECUTAR()
                Call CARGAR_DATAGRIDVIEW()
                Call ARMAR_DATAGRIDVIEW()
            End If
        Else
            MsgBox("Debe Ingresar un Codigo de Colaborador válido", vbInformation, "Mensaje del Sistema")
            Me.TextBox1.SelectAll()
            Me.TextBox1.Focus()
            Exit Sub
        End If
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter("Select * from [CONSULTA TIEMPO EXTRAORDINARIO POR COLABORADOR] WHERE ID_USUARIO = " & isuID_USUARIO & " ORDER BY FECHA", Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[CONSULTA TIEMPO EXTRAORDINARIO POR COLABORADOR]")
        DGV1.DataSource = MiDataSet.Tables("[CONSULTA TIEMPO EXTRAORDINARIO POR COLABORADOR]")
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW()
        Me.DGV1.Columns(0).HeaderText = "Fecha"
        Me.DGV1.Columns(0).Width = 80
        Me.DGV1.Columns(0).Visible = True
        Me.DGV1.Columns(0).DefaultCellStyle.Format = "d"
        Me.DGV1.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(1).HeaderText = "Dia"
        Me.DGV1.Columns(1).Width = 100
        Me.DGV1.Columns(1).Visible = True
        Me.DGV1.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(2).HeaderText = "Situacion"
        Me.DGV1.Columns(2).Width = 120
        Me.DGV1.Columns(2).Visible = True
        Me.DGV1.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(3).HeaderText = "Horario de Entrada"
        Me.DGV1.Columns(3).Width = 100
        Me.DGV1.Columns(3).Visible = True
        Me.DGV1.Columns(3).DefaultCellStyle.Format = "t"
        Me.DGV1.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(4).HeaderText = "Hora de Marca de Entrada"
        Me.DGV1.Columns(4).Width = 100
        Me.DGV1.Columns(4).Visible = True
        Me.DGV1.Columns(4).DefaultCellStyle.Format = "t"
        Me.DGV1.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(5).HeaderText = "Horario de Salida"
        Me.DGV1.Columns(5).Width = 100
        Me.DGV1.Columns(5).Visible = True
        Me.DGV1.Columns(5).DefaultCellStyle.Format = "t"
        Me.DGV1.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV1.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(6).HeaderText = "Hora de Marca de Salida"
        Me.DGV1.Columns(6).Width = 100
        Me.DGV1.Columns(6).Visible = True
        Me.DGV1.Columns(6).DefaultCellStyle.Format = "t"
        Me.DGV1.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV1.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(7).HeaderText = "Total Horas por dia"
        Me.DGV1.Columns(7).Width = 100
        Me.DGV1.Columns(7).Visible = True
        'Me.DGV1.Columns(7).DefaultCellStyle.Format = "t"
        Me.DGV1.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV1.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV1.Columns(7).DefaultCellStyle.BackColor = Color.LightGreen

        Me.DGV1.Columns(8).HeaderText = "ID_USUARIO"
        Me.DGV1.Columns(8).Width = 0
        Me.DGV1.Columns(8).Visible = False
    End Sub
    Private Sub DGV1_RowPrePaint(sender As Object, e As DataGridViewRowPrePaintEventArgs) Handles DGV1.RowPrePaint
        Dim style As New DataGridViewCellStyle
        style.Font = New Font(DGV1.Font, FontStyle.Bold)
        For Each Row As DataGridViewRow In DGV1.Rows
            If Row.Cells(2).Value = "AUSENCIA INJUSTIFICADA" Or Row.Cells(2).Value = "OMISION DE MARCA" Then
                Row.DefaultCellStyle.ForeColor = Color.Red
            End If
            If Row.Cells(2).Value = "VACACIONES" Then
                Row.DefaultCellStyle.ForeColor = Color.Blue
            End If
            If Row.Cells(2).Value = "SUBSIDIO" Then
                Row.DefaultCellStyle.ForeColor = Color.DarkOrange
            End If
            If Row.Cells(2).Value = "AJUSTE MENSUAL (+)" Then
                Row.DefaultCellStyle.ForeColor = Color.Red
            End If
            If Row.Cells(2).Value = "AJUSTE MENSUAL (-)" Then
                Row.DefaultCellStyle.ForeColor = Color.Red
            End If
            'If Row.Cells(1).Value = "SÁBADO" Or Row.Cells(1).Value = "DOMINGO" Then
            '    Row.DefaultCellStyle.ForeColor = Color.Red
            'End If
        Next
    End Sub
    Private Sub BORRAR_DATOS_USUARIO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "DELETE FROM [CONSULTA TIEMPO EXTRAORDINARIO POR COLABORADOR] WHERE ID_USUARIO = " & CInt(isuID_USUARIO) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim HAY_HORARIO As Boolean
    Dim vID_T_H As Integer
    Private Sub BUSCAR_HORARIO()
        Me.Cursor = Cursors.WaitCursor
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT ID_T_HORARIO FROM [VISTA MAESTRO DE PERSONAS] WHERE CODIGO = '" & Me.TextBox1.Text & "' AND ID_ESTADO_P = 1", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                Me.Cursor = Cursors.WaitCursor
                HAY_HORARIO = True
                vID_T_H = MiDataTable.Rows(0).Item(0).ToString
                Call BUSCAR_HORAS()
            Else
                HAY_HORARIO = False
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
        Me.Cursor = Cursors.Default
    End Sub
    Dim HORARIO1, HORARIO2 As String
    Private Sub BUSCAR_HORAS()
        Me.Cursor = Cursors.WaitCursor
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE HORARIO] WHERE ID_T_HORARIO = " & vID_T_H & "", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                HORARIO1 = MiDataTable.Rows(0).Item(2).ToString
                HORARIO2 = MiDataTable.Rows(0).Item(3).ToString
            Else
                HORARIO1 = ""
                HORARIO2 = ""
                HAY_HORARIO = False
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
        Me.Cursor = Cursors.Default
    End Sub
    Dim vFECHA As String
    Dim vDIA As String
    Dim vSITUACION As String
    Dim vH1, vH2 As System.DateTime
    Private Sub EJECUTAR()
        Call BUSCAR_HORARIO()
        If HAY_HORARIO = False Then
            MsgBox("El Colaborador seleccionado no tiene asignado un Horario, por favor asiganr Horario en ventana de Expedientes Activos", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        FECHA01 = Mid(Me.DateTimePicker1.Text, 1, 10)
        FECHA1 = Mid(FECHA01, 1, 10)
        FECHA02 = Mid(Me.DateTimePicker2.Text, 1, 10)
        FECHA2 = Mid(FECHA02, 1, 10)
        CADENAsql1 = "SET LANGUAGE ESPAÑOL Select ID, NIVEL1, NIVEL2, NIVEL3, NIVEL4, NIVEL5, ORD, ID_M_A, ID_M_P, CODIGO, UPPER(DATENAME (dw, DIA)) AS DIALETRAS, DIA, ID_T_SIT_P, D_T_SIT_P, ID_SIT_P, D_SIT_P, SIGLA, VALOR_SIT, VALOR_DIA, OBSERVACIONES, DETALLE_1, DETALLE_2, DETALLE_3, D_ADMIN, USUARIO_ACT, FECHA_ACT from [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE CODIGO = '" & Me.TextBox1.Text & "' AND DIA BETWEEN '" & FECHA1 & "' AND '" & FECHA2 & "' AND (D_SIT_P <> 'AJUSTE MENSUAL (+)' AND D_SIT_P <> 'AJUSTE MENSUAL (-)') ORDER BY CODIGO, DIA SET LANGUAGE ENGLISH"
        Me.Cursor = Cursors.WaitCursor
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter(CADENAsql1, Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                For Y = 0 To MiDataTable.Rows.Count - 1
                    Application.DoEvents()
                    Me.Cursor = Cursors.WaitCursor
                    vFECHA = Mid(MiDataTable.Rows(Y).Item(11).ToString, 1, 10)
                    vDIA = MiDataTable.Rows(Y).Item(10).ToString
                    vSITUACION = MiDataTable.Rows(Y).Item(15).ToString
                    Call BUSCAR_HORA_I()
                    Call BUSCAR_HORA_F()
                    Call AGREGAR_REGISTRO()
                    Me.Cursor = Cursors.Default
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
    Dim fHORA_MARCA1 As String
    Dim fHORA_MARCA2 As String
    Dim CH As String

    Dim FECHADESDE As String
    Dim FECHAHASTA As String
    Private Sub AGREGAR_REGISTRO()
        Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(vFECHA, 1, 10) + "', 105), 23)"

        If vH1 = "1900-01-01 12:00:00.000" Then
            fHORA_MARCA1 = ""
        Else
            fHORA_MARCA1 = vH1.Hour & ":" & vH1.Minute & ":" & vH1.Second
        End If

        If vH2 = "1900-01-01 12:00:00.000" Then
            fHORA_MARCA2 = ""
        Else
            fHORA_MARCA2 = vH2.Hour & ":" & vH2.Minute & ":" & vH2.Second
        End If

        If fHORA_MARCA1 <> "" And fHORA_MARCA2 <> "" Then
            Dim final As String = fHORA_MARCA2
            Dim inicial As String = fHORA_MARCA1
            Dim tiempoI As TimeSpan = TimeSpan.Parse(inicial)
            Dim tiempoF As TimeSpan = TimeSpan.Parse(final)
            Dim resta As TimeSpan = tiempoF - tiempoI
            CH = resta.ToString
        Else
            CH = ""
        End If

        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [CONSULTA TIEMPO EXTRAORDINARIO POR COLABORADOR] (FECHA, DIA, SITUACION, HORA_I_H, HORA_I, HORA_F_H, HORA_F, TOT_HL_X_DIA, ID_USUARIO)" &
                              "values (" & F1 & ", '" & vDIA & "', '" & vSITUACION & "', '" & HORARIO1 & "', '" & fHORA_MARCA1 & "', '" & HORARIO2 & "', '" & fHORA_MARCA2 & "', '" & CH & "'," & isuID_USUARIO & ")"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            TITULO_EXCEL = "TIEMPO EXTRAORDINARIO"
            EXPORTAR_DATOS_A_EXCEL(DGV1, "Colaborador: " & Me.TextBox1.Text & " - " & Me.TextBox2.Text & "[" & Mid(Me.DateTimePicker1.Value, 1, 10) & " - " & Mid(Me.DateTimePicker2.Value, 1, 10))
        Catch ex As Exception
            MessageBox.Show("No se puede generar el documento Excel.", "Mensaje del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub BUSCAR_HORA_I()
        '-------------*************----------------------------------------
        FECHA01 = "Convert(VARCHAR(10), Convert(Date, '" + Mid(vFECHA, 1, 10) + "', 105), 23)"
        '-------------*************----------------------------------------
        CADENAsql = "SET LANGUAGE Español SELECT [IDEMPRESA], [RELOJ], [ID_EMPLEADO], DATENAME (dw, DIA) AS DIAL, [FECHA_MARCA], [HORA_MARCA], [FECHA_CARGA], [ARCHIVO], [NUMERO], [PROGRAMADO], [TIPO_MARCA], [OBSERVACION] FROM [dbo].[VISTA MAESTRO DE MARCAS] WHERE (ID_EMPLEADO = '" & Me.TextBox1.Text & "') AND (FECHA_MARCA = " & FECHA01 & ")  AND TIPO_MARCA = 'ENTRADA' ORDER BY HORA_MARCA DESC SET LANGUAGE ENGLISH"
        Me.Cursor = Cursors.WaitCursor
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
                Application.DoEvents()
                Me.Cursor = Cursors.WaitCursor
                vH1 = MiDataTable.Rows(0).Item(5).ToString
                FECHADESDE = vH1
                Me.Cursor = Cursors.Default
            Else
                vH1 = "1900-01-01 12:00:00.000"
                FECHADESDE = "1900-01-01 12:00:00.000"
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
    Private Sub BUSCAR_HORA_F()
        '-------------*************----------------------------------------
        FECHA01 = "Convert(VARCHAR(10), Convert(Date, '" + Mid(vFECHA, 1, 10) + "', 105), 23)"
        '-------------*************----------------------------------------
        CADENAsql = "SET LANGUAGE Español SELECT [IDEMPRESA], [RELOJ], [ID_EMPLEADO], DATENAME (dw, DIA) AS DIAL, [FECHA_MARCA], [HORA_MARCA], [FECHA_CARGA], [ARCHIVO], [NUMERO], [PROGRAMADO], [TIPO_MARCA], [OBSERVACION] FROM [dbo].[VISTA MAESTRO DE MARCAS] WHERE (ID_EMPLEADO = '" & Me.TextBox1.Text & "') AND (FECHA_MARCA = " & FECHA01 & ")  AND TIPO_MARCA = 'SALIDA' ORDER BY HORA_MARCA DESC SET LANGUAGE ENGLISH"
        Me.Cursor = Cursors.WaitCursor
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
                Application.DoEvents()
                Me.Cursor = Cursors.WaitCursor
                vH2 = MiDataTable.Rows(0).Item(5).ToString
                FECHAHASTA = vH2
                Me.Cursor = Cursors.Default
            Else
                vH2 = "1900-01-01 12:00:00.000"
                FECHAHASTA = "1900-01-01 12:00:00.000"
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
End Class