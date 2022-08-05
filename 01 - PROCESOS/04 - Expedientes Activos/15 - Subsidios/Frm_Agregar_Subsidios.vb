Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Agregar_Subsidios
    Dim ANTxDIA As Double
    Dim valTSit As Integer
    Dim valSit As Integer
    Dim valValor As String
    Private Sub Frm_Agregar_Subsidios_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = "0"
            Me.TextBox2.Text = ""
            Me.TextBox3.Text = ""
            Me.TextBox4.Text = "0"
            Me.TextBox5.Text = "0"
            Me.TextBox6.Text = ""
            Me.TextBox7.Text = ""
            Me.TextBox8.Text = ""
            Me.TextBox9.Text = ""
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Agregar_Subsidios_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = "0"
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_SUBSIDIOS()
        Me.TextBox4.Text = "0"
        Me.TextBox5.Text = "0"
        Me.TextBox6.Text = ""
        Me.TextBox7.Text = ""
        Call CARGAR_COMBO_CAT_DE_ESPECIALIDADES_MEDICAS()
        Me.TextBox8.Text = ""
        Me.TextBox9.Text = ""
        Me.TextBox2.Focus()
    End Sub
    Private Sub Button7_Click(sender As Object, e As System.EventArgs) Handles Button7.Click
        Me.TextBox1.Text = "0"
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = "0"
        Me.TextBox5.Text = "0"
        Me.TextBox6.Text = ""
        Me.TextBox7.Text = ""
        Me.TextBox8.Text = ""
        Me.TextBox9.Text = ""
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESPECIALIDADES_MEDICAS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESPECIALIDADES MEDICAS] WHERE ID_T_ESTUDIO = 15  order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox2
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_E_MED"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_SUBSIDIOS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE SUBSIDIOS] order by ID_T_SUBS", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_T_SUBS"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
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
    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
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
    Private Sub DateTimePicker2_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker2.KeyDown
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
    Private Sub DateTimePicker3_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub DateTimePicker4_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker4.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub DateTimePicker5_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker5.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub DateTimePicker6_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker6.KeyDown
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
    Private Sub TextBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox7.KeyDown
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
    Private Sub ComboBox3_KeyDown(sender As Object, e As KeyEventArgs)
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
    Private Sub TextBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox9.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub GENERAR_CODIGO()
        CODIGO = 0
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            MiDataAdapter = New SqlDataAdapter("SELECT ID_SUBS FROM [MAESTRO DE SUBSIDIOS] ORDER BY ID_SUBS", Conexion.CxRRHH)
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.TextBox1.Text = "0"
        Call GENERAR_CODIGO()
        If TextBox1.Text = "0" Or TextBox1.Text = "" Then
            MsgBox("No se ha generado el Codigo del registro", vbInformation, "Mensaje del Sistema")
            Me.TextBox1.Text = "0"
            Me.Button7.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox2.Text) = 0 Then
            Me.TextBox2.Text = "0"
        End If
        If Len(Me.TextBox3.Text) = 0 Then
            Me.TextBox3.Text = "0"
        End If
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Tipo de Subsidio Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox4.Text) = 0 Then
            Me.TextBox4.Text = "0"
        End If
        If Len(Me.TextBox5.Text) = 0 Then
            Me.TextBox5.Text = "0"
        End If
        If Len(Me.TextBox6.Text) = 0 Then
            Me.TextBox6.Text = "."
        End If
        If Len(Me.TextBox7.Text) = 0 Then
            Me.TextBox7.Text = "0"
        End If
        If Len(Me.ComboBox2.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar la Especialidad Medica Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox8.Text) = 0 Then
            Me.TextBox8.Text = "."
        End If
        If Len(Me.TextBox9.Text) = 0 Then
            Me.TextBox9.Text = "."
        End If
        Call GRABAR_REGISTRO()
        vDET1 = "EXPEDIENTE: " & Me.TextBox2.Text & "; NO. BOLETA: " & Me.TextBox3.Text & "; FECHAS: " & Mid(Me.DateTimePicker1.Value, 1, 10) & " - " & Mid(Me.DateTimePicker2.Value, 1, 10)
        vDET2 = "MEDICO: (" & Me.TextBox7.Text & ") " & Me.TextBox6.Text
        vDET3 = "DIAGNOSTICO: " & Me.TextBox8.Text
        Call GRABAR_SITUACION()
        vmssID_SUBS = Me.TextBox1.Text
        CADENAsql = "SELECT * FROM [VISTA MAESTRO DE SUBSIDIOS] WHERE ID_M_P = " & vID_M_P & "ORDER BY FECHA_I"
        Call CARGAR_DATAGRIDVIEW_17()
        Call ARMAR_DATAGRIDVIEW_17()
        Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_17()
    End Sub
    Private Sub GRABAR_REGISTRO()
        Dim fFECHA_I As String
        If Me.DateTimePicker1.Checked = True Then
            fFECHA_I = Mid(Me.DateTimePicker1.Value, 1, 10)
        Else
            fFECHA_I = "NULL"
        End If
        Dim f1 As Object = If(fFECHA_I <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_I + "', 105), 23)", "NULL")

        Dim fFECHA_F As String
        If Me.DateTimePicker2.Checked = True Then
            fFECHA_F = Mid(Me.DateTimePicker2.Value, 1, 10)
        Else
            fFECHA_F = "NULL"
        End If
        Dim f2 As Object = If(fFECHA_F <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_F + "', 105), 23)", "NULL")

        Dim fFECHA_PARTO As String
        If Me.DateTimePicker3.Checked = True Then
            fFECHA_PARTO = Mid(Me.DateTimePicker3.Value, 1, 10)
        Else
            fFECHA_PARTO = "NULL"
        End If
        Dim f3 As Object = If(fFECHA_PARTO <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_PARTO + "', 105), 23)", "NULL")

        Dim fFECHA_PARTO_PROB As String
        If Me.DateTimePicker4.Checked = True Then
            fFECHA_PARTO_PROB = Mid(Me.DateTimePicker4.Value, 1, 10)
        Else
            fFECHA_PARTO_PROB = "NULL"
        End If
        Dim f4 As Object = If(fFECHA_PARTO_PROB <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_PARTO_PROB + "', 105), 23)", "NULL")

        Dim fFECHA_ACC_LAB As String
        If Me.DateTimePicker5.Checked = True Then
            fFECHA_ACC_LAB = Mid(Me.DateTimePicker5.Value, 1, 10)
        Else
            fFECHA_ACC_LAB = "NULL"
        End If
        Dim f5 As Object = If(fFECHA_ACC_LAB <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_ACC_LAB + "', 105), 23)", "NULL")

        Dim fFECHA_DECLARACION As String
        If Me.DateTimePicker6.Checked = True Then
            fFECHA_DECLARACION = Mid(Me.DateTimePicker6.Value, 1, 10)
        Else
            fFECHA_DECLARACION = "NULL"
        End If
        Dim f6 As Object = If(fFECHA_DECLARACION <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_DECLARACION + "', 105), 23)", "NULL")

        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()

        Dim SQL As String
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            SQL = "INSERT INTO [MAESTRO DE SUBSIDIOS] (ID_SUBS, ID_M_P, NO_EXPEDIENTE, NO_BOLETA, ID_T_SUBS, N_ORDEN, FECHA_I, FECHA_F, CANT_DIAS, FECHA_PARTO, FECHA_PARTO_PROB, FECHA_ACC_LAB, FECHA_DECLARACION, VALOR_X_DIA, VALOR_TOTAL, MEDICO, N_MINSA, ID_E_MED, DIAGNOSTICO, OBSERVACIONES)" &
                   "values (" & Me.TextBox1.Text & ", " & vID_M_P & ", '" & Me.TextBox2.Text & "', '" & Me.TextBox3.Text & "', " & Me.ComboBox1.SelectedValue & ", '" & Me.TextBox4.Text & "', " & f1 & ", " & f2 & ", " & Me.TextBox5.Text & ", " & f3 & ", " & f4 & ", " & f5 & ", " & f6 & ", 0, 0, '" & Me.TextBox6.Text & "', '" & Me.TextBox7.Text & "', " & Me.ComboBox2.SelectedValue & ", '" & Me.TextBox8.Text & "', '" & Me.TextBox9.Text & "')"
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
    Private Sub CARGAR_DATAGRIDVIEW_17()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE SUBSIDIOS]")
        Frm_Expedientes_Activos.DGV17.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE SUBSIDIOS]")
        Frm_Expedientes_Activos.Label103.Text = Frm_Expedientes_Activos.DGV17.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_17()
        Frm_Expedientes_Activos.DGV17.Columns(0).HeaderText = "Id"
        Frm_Expedientes_Activos.DGV17.Columns(0).Width = 30
        Frm_Expedientes_Activos.DGV17.Columns(0).Visible = True
        Frm_Expedientes_Activos.DGV17.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV17.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Frm_Expedientes_Activos.DGV17.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Frm_Expedientes_Activos.DGV17.Columns(1).HeaderText = "ID_M_P"
        Frm_Expedientes_Activos.DGV17.Columns(1).Width = 0
        Frm_Expedientes_Activos.DGV17.Columns(1).Visible = False

        Frm_Expedientes_Activos.DGV17.Columns(2).HeaderText = "No. Exp."
        Frm_Expedientes_Activos.DGV17.Columns(2).Width = 80
        Frm_Expedientes_Activos.DGV17.Columns(2).Visible = True
        Frm_Expedientes_Activos.DGV17.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV17.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV17.Columns(3).HeaderText = "Boleta"
        Frm_Expedientes_Activos.DGV17.Columns(3).Width = 80
        Frm_Expedientes_Activos.DGV17.Columns(3).Visible = True
        Frm_Expedientes_Activos.DGV17.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV17.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV17.Columns(3).Frozen = True

        Frm_Expedientes_Activos.DGV17.Columns(4).HeaderText = "ID_T_SUBS"
        Frm_Expedientes_Activos.DGV17.Columns(4).Width = 0
        Frm_Expedientes_Activos.DGV17.Columns(4).Visible = False

        Frm_Expedientes_Activos.DGV17.Columns(5).HeaderText = "Tipo"
        Frm_Expedientes_Activos.DGV17.Columns(5).Width = 220
        Frm_Expedientes_Activos.DGV17.Columns(5).Visible = True
        Frm_Expedientes_Activos.DGV17.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV17.Columns(6).HeaderText = "Ord"
        Frm_Expedientes_Activos.DGV17.Columns(6).Width = 40
        Frm_Expedientes_Activos.DGV17.Columns(6).Visible = True
        Frm_Expedientes_Activos.DGV17.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV17.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Frm_Expedientes_Activos.DGV17.Columns(7).HeaderText = "Fecha Inicio"
        Frm_Expedientes_Activos.DGV17.Columns(7).Width = 90
        Frm_Expedientes_Activos.DGV17.Columns(7).Visible = True
        Frm_Expedientes_Activos.DGV17.Columns(7).DefaultCellStyle.Format = "d"
        Frm_Expedientes_Activos.DGV17.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV17.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV17.Columns(8).HeaderText = "Fecha Fin"
        Frm_Expedientes_Activos.DGV17.Columns(8).Width = 90
        Frm_Expedientes_Activos.DGV17.Columns(8).Visible = True
        Frm_Expedientes_Activos.DGV17.Columns(8).DefaultCellStyle.Format = "d"
        Frm_Expedientes_Activos.DGV17.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV17.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV17.Columns(9).HeaderText = "Dias"
        Frm_Expedientes_Activos.DGV17.Columns(9).Width = 40
        Frm_Expedientes_Activos.DGV17.Columns(9).Visible = True
        Frm_Expedientes_Activos.DGV17.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV17.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Frm_Expedientes_Activos.DGV17.Columns(10).HeaderText = "Fecha Parto"
        Frm_Expedientes_Activos.DGV17.Columns(10).Width = 90
        Frm_Expedientes_Activos.DGV17.Columns(10).Visible = True
        Frm_Expedientes_Activos.DGV17.Columns(10).DefaultCellStyle.Format = "d"
        Frm_Expedientes_Activos.DGV17.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV17.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV17.Columns(11).HeaderText = "Fecha Parto Prob."
        Frm_Expedientes_Activos.DGV17.Columns(11).Width = 90
        Frm_Expedientes_Activos.DGV17.Columns(11).Visible = True
        Frm_Expedientes_Activos.DGV17.Columns(11).DefaultCellStyle.Format = "d"
        Frm_Expedientes_Activos.DGV17.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV17.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV17.Columns(12).HeaderText = "Fecha Acc. Lab."
        Frm_Expedientes_Activos.DGV17.Columns(12).Width = 90
        Frm_Expedientes_Activos.DGV17.Columns(12).Visible = True
        Frm_Expedientes_Activos.DGV17.Columns(12).DefaultCellStyle.Format = "d"
        Frm_Expedientes_Activos.DGV17.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV17.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV17.Columns(13).HeaderText = "Fecha Declaracion"
        Frm_Expedientes_Activos.DGV17.Columns(13).Width = 90
        Frm_Expedientes_Activos.DGV17.Columns(13).Visible = True
        Frm_Expedientes_Activos.DGV17.Columns(13).DefaultCellStyle.Format = "d"
        Frm_Expedientes_Activos.DGV17.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV17.Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV17.Columns(14).HeaderText = "VALOR_X_DIA"
        Frm_Expedientes_Activos.DGV17.Columns(14).Width = 0
        Frm_Expedientes_Activos.DGV17.Columns(14).Visible = False

        Frm_Expedientes_Activos.DGV17.Columns(15).HeaderText = "VALOR_X_DIA"
        Frm_Expedientes_Activos.DGV17.Columns(15).Width = 0
        Frm_Expedientes_Activos.DGV17.Columns(15).Visible = False

        Frm_Expedientes_Activos.DGV17.Columns(16).HeaderText = "Medico"
        Frm_Expedientes_Activos.DGV17.Columns(16).Width = 230
        Frm_Expedientes_Activos.DGV17.Columns(16).Visible = True
        Frm_Expedientes_Activos.DGV17.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV17.Columns(17).HeaderText = "No. MINSA"
        Frm_Expedientes_Activos.DGV17.Columns(17).Width = 60
        Frm_Expedientes_Activos.DGV17.Columns(17).Visible = True
        Frm_Expedientes_Activos.DGV17.Columns(17).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV17.Columns(17).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV17.Columns(18).HeaderText = "ID_E_MED"
        Frm_Expedientes_Activos.DGV17.Columns(18).Width = 0
        Frm_Expedientes_Activos.DGV17.Columns(18).Visible = False

        Frm_Expedientes_Activos.DGV17.Columns(19).HeaderText = "Especialidad"
        Frm_Expedientes_Activos.DGV17.Columns(19).Width = 0
        Frm_Expedientes_Activos.DGV17.Columns(19).Visible = False
        Frm_Expedientes_Activos.DGV17.Columns(19).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV17.Columns(20).HeaderText = "Diagnostico"
        Frm_Expedientes_Activos.DGV17.Columns(20).Width = 500
        Frm_Expedientes_Activos.DGV17.Columns(20).Visible = True
        Frm_Expedientes_Activos.DGV17.Columns(20).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV17.Columns(21).HeaderText = "Observaciones"
        Frm_Expedientes_Activos.DGV17.Columns(21).Width = 100
        Frm_Expedientes_Activos.DGV17.Columns(21).Visible = True
        Frm_Expedientes_Activos.DGV17.Columns(21).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_17()
        Dim I As Integer
        For I = 0 To Frm_Expedientes_Activos.DGV17.RowCount - 1
            If Frm_Expedientes_Activos.DGV17.Item(0, I).Value = vmssID_SUBS Then
                Frm_Expedientes_Activos.DGV17.Rows(I).Selected = True
                Frm_Expedientes_Activos.DGV17.CurrentCell = Frm_Expedientes_Activos.DGV17.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub GRABAR_SITUACION()
        Me.Cursor = Cursors.WaitCursor
        Me.DateTimePicker0.Value = Me.DateTimePicker1.Value
        Application.DoEvents()
        For I2 As Integer = 1 To DateDiff(DateInterval.Day, Me.DateTimePicker1.Value, Me.DateTimePicker2.Value.AddDays(1))
            Me.DateTimePicker0.Refresh()
            '-------------- SITUACIONES
            Call BUSCAR_SI_EXISTE_DIA_EN_MAESTRO_DE_SITUACION_DE_PERSONAL()
            If EXISTE_DIA_EN_MAESTRO_DE_SITUACION_DE_PERSONAL = True Then
                Call UPD_ROW()
            Else
                Call ADD_ROW()
            End If
            '-------------- SUBSIDIOS
            Call BUSCAR_SI_EXISTE_DIA_EN_MAESTRO_DE_SUBSIDIOS_DETALLES_DIAS()
            If EXISTE_DIA_EN_MAESTRO_DE_SUBSIDIOS_DETALLES_DIAS = True Then
                Call UPD_ROW_S()
            Else
                Call ADD_ROW_S()
            End If
            Me.DateTimePicker0.Value = Me.DateTimePicker0.Value.AddDays(1)
            Me.DateTimePicker0.Refresh()
        Next
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub UPD_ROW()
        Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker0.Value, 1, 10) + "', 105), 23)"
        Dim CD As Integer = 30
        If CD = 30 Then
            ANTxDIA = 0.0833333333333333
        End If
        valTSit = 2
        valSit = 14
        valValor = 0
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE SITUACION DE PERSONAL] SET ID_T_SIT_P = " & valTSit & ", ID_SIT_P = " & valSit & ", VALOR_SIT = '" & CDbl(valValor) & "', VALOR_DIA = '" & CDbl(ANTxDIA) & "', OBSERVACIONES = 'INGRESADO DESDE EL MODULO DE EXPEDIENTES', DETALLE_1 = '" & vDET1 & "', DETALLE_2 = '" & vDET2 & "', DETALLE_3 = '" & vDET3 & "' WHERE DIA = " & F1 & " AND ID_M_P = " & Frm_Expedientes_Activos.TextBox35.Text & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub UPD_ROW_S()

    End Sub
    Private Sub ADD_ROW_S()
        Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker0.Value, 1, 10) + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE SUBSIDIOS DETALLES DIAS] (ID_SUBS, CODIGO, NO_EXPEDIENTE, NO_BOLETA, ID_T_SUBS, FECHA_S, DIAGNOSTICO)" &
                              "values (" & Me.TextBox1.Text & ", '" & Frm_Expedientes_Activos.TextBox5.Text & "', '" & Me.TextBox2.Text & "', '" & Me.TextBox3.Text & "', " & Me.ComboBox1.SelectedValue & ", " & F1 & ", '" & Trim(Me.TextBox8.Text) & "')"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ADD_ROW()
        Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker0.Value, 1, 10) + "', 105), 23)"
        Dim CD As Integer = 30
        If CD = 30 Then
            ANTxDIA = 0.0833333333333333
        End If
        valTSit = 2
        valSit = 14
        valValor = 0
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE SITUACION DE PERSONAL] (ID_M_P, DIA, ID_T_SIT_P, ID_SIT_P, VALOR_SIT, VALOR_DIA, OBSERVACIONES, DETALLE_1, DETALLE_2, DETALLE_3)" &
                              "values (" & CInt(Frm_Expedientes_Activos.TextBox35.Text) & ", " & F1 & ", " & valTSit & ", " & valSit & ", '" & CDbl(valValor) & "', '" & CDbl(ANTxDIA) & "','INGRESADO DESDE EL MODULO DE EXPEDIENTES', '" & vDET1 & "', '" & vDET2 & "', '" & vDET3 & "')"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim EXISTE_DIA_EN_MAESTRO_DE_SITUACION_DE_PERSONAL As Boolean
    Private Sub BUSCAR_SI_EXISTE_DIA_EN_MAESTRO_DE_SITUACION_DE_PERSONAL()
        cEMPLEADO = Format(CInt(Frm_Expedientes_Activos.TextBox5.Text), "0000000000")
        Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker0.Value, 1, 10) + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable3 As New DataTable
        Dim MiDataAdapter3 As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter3 = New SqlDataAdapter("SELECT * FROM [VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE (CODIGO = '" & cEMPLEADO & "') AND (DIA = " & F1 & ") ORDER BY  CODIGO, DIA", Conexion.CxRRHH)
            MiDataTable3.Clear()
            MiDataAdapter3.Fill(MiDataTable3)
            If MiDataTable3.Rows.Count <> 0 Then
                EXISTE_DIA_EN_MAESTRO_DE_SITUACION_DE_PERSONAL = True
            Else
                EXISTE_DIA_EN_MAESTRO_DE_SITUACION_DE_PERSONAL = False
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim EXISTE_DIA_EN_MAESTRO_DE_SUBSIDIOS_DETALLES_DIAS As Boolean
    Private Sub BUSCAR_SI_EXISTE_DIA_EN_MAESTRO_DE_SUBSIDIOS_DETALLES_DIAS()
        cEMPLEADO = Format(CInt(Frm_Expedientes_Activos.TextBox5.Text), "0000000000")
        Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker0.Value, 1, 10) + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable3 As New DataTable
        Dim MiDataAdapter3 As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter3 = New SqlDataAdapter("SELECT * FROM [MAESTRO DE SUBSIDIOS DETALLES DIAS] WHERE (CODIGO = '" & cEMPLEADO & "') AND (FECHA_S = " & F1 & ") ORDER BY  CODIGO, FECHA_S", Conexion.CxRRHH)
            MiDataTable3.Clear()
            MiDataAdapter3.Fill(MiDataTable3)
            If MiDataTable3.Rows.Count <> 0 Then
                EXISTE_DIA_EN_MAESTRO_DE_SUBSIDIOS_DETALLES_DIAS = True
            Else
                EXISTE_DIA_EN_MAESTRO_DE_SUBSIDIOS_DETALLES_DIAS = False
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