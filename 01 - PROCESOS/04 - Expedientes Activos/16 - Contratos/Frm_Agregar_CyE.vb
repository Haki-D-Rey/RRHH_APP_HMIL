Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Agregar_CyE
    Private Sub Frm_Agregar_CyE_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = "0"
            Me.TextBox3.Text = ""
            Me.TextBox2.Text = ""
            Me.ComboBox1.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Agregar_CyE_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = "0"
        Call CARGAR_COMBO_CAT_DE_CONTRATOS()
        Me.DateTimePicker1.Value = Now
        Me.DateTimePicker2.Value = Now
        Me.DateTimePicker3.Value = Now
        Me.TextBox3.Text = ""
        Call CARGAR_COMBO_CAT_DE_CLASIFICACION_DE_EVALUACIONES()
        Me.TextBox2.Text = ""
        Me.ComboBox1.Focus()
    End Sub
    Private Sub Button7_Click(sender As Object, e As System.EventArgs) Handles Button7.Click
        Me.TextBox1.Text = "0"
        Me.TextBox3.Text = ""
        Me.TextBox2.Text = ""
        Me.ComboBox1.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_CLASIFICACION_DE_EVALUACIONES()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE CLASIFICACION DE EVALUACIONES] order by ID_C_E", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox2
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_C_E"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_CONTRATOS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE CONTRATOS] order by ID_CONT", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_CONT"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
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
    Private Sub CheckBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox1.KeyDown
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
    Private Sub GENERAR_CODIGO()
        CODIGO = 0
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            MiDataAdapter = New SqlDataAdapter("SELECT ID_M_CYE FROM [MAESTRO DE CONTRATOS Y EVALUACIONES] ORDER BY ID_M_CYE", Conexion.CxRRHH)
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
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Tipo Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox2.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar la Clasificacion Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox2.Text) = 0 Then
            Me.TextBox2.Text = ""
        End If
        If Me.TextBox3.Text = "" Then
            Me.TextBox3.Text = ""
        End If
        If Me.ComboBox1.Text = "1° CONTRATO" Or Me.ComboBox1.Text = "2° CONTRATO" Or Me.ComboBox1.Text = "CONTRATO INDEFINIDO" Then
            Me.TextBox3.Text = ""
            Me.ComboBox2.Text = "SIN DEFINIR"
            GoTo SALTO
        End If
        If Not IsNumeric(Me.TextBox3.Text) Then
            Me.TextBox3.Text = "0"
        End If
        If IsNumeric(CInt(Me.TextBox3.Text)) Then
            Dim vNOTA As Double
            vNOTA = CDbl(Me.TextBox3.Text)
            If vNOTA < 60 Then
                Me.ComboBox2.Text = "DEFICIENTE"
            End If
            If vNOTA > 59 And vNOTA < 69 Then
                Me.ComboBox2.Text = "REGULAR"
            End If
            If vNOTA > 68 And vNOTA < 79 Then
                Me.ComboBox2.Text = "BUENO"
            End If
            If vNOTA > 78 And vNOTA < 89 Then
                Me.ComboBox2.Text = "MUY BUENO"
            End If
            If vNOTA > 88 And vNOTA < 101 Then
                Me.ComboBox2.Text = "EXCELENTE"
            End If
        Else
            MsgBox("El dato ingresado no es númerico o no es váildo", vbInformation, "Mensaje del Sistema")
            Me.TextBox3.SelectAll()
            Me.TextBox3.Focus()
            Exit Sub
        End If
SALTO:
        Call GRABAR_REGISTRO()
        mcyeID_M_CYE = Me.TextBox1.Text
        CADENAsql = "SELECT * FROM [dbo].[VISTA MAESTRO DE CONTRATOS Y EVALUACIONES] WHERE ID_M_P = " & vID_M_P & " ORDER BY [CODIGO], [ID_CONT]"
        Call CARGAR_DATAGRIDVIEW_18()
        Call ARMAR_DATAGRIDVIEW_18()
        Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_18()
        Me.TextBox1.Text = "0"
        Me.TextBox3.Text = ""
        Me.TextBox2.Text = ""
        Me.ComboBox1.Focus()
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


        Dim fFECHA_E As String
        If Me.DateTimePicker3.Checked = True Then
            fFECHA_E = Mid(Me.DateTimePicker3.Value, 1, 10)
        Else
            fFECHA_E = "NULL"
        End If
        Dim f3 As Object = If(fFECHA_E <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_E + "', 105), 23)", "NULL")


        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()

        Dim SQL As String
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            SQL = "INSERT INTO [MAESTRO DE CONTRATOS Y EVALUACIONES] (ID_M_CYE, ID_M_P, ID_CONT, FECHA_I, FECHA_F, FIRMADO, FECHA_EVA, PROMEDIO_NOTA, ID_C_E, OBSERVACIONES)" &
                   "values (" & Me.TextBox1.Text & ", " & vID_M_P & ", " & Me.ComboBox1.SelectedValue & ", " & f1 & ", " & f2 & ", '" & Me.CheckBox1.Checked & "', " & f3 & ", '" & Me.TextBox3.Text & "', " & Me.ComboBox2.SelectedValue & ", '" & Me.TextBox2.Text & "')"
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
    Private Sub CARGAR_DATAGRIDVIEW_18()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE CONTRATOS Y EVALUACIONES]")
        Frm_Expedientes_Activos.DGV18.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE CONTRATOS Y EVALUACIONES]")
        Frm_Expedientes_Activos.Label109.Text = Frm_Expedientes_Activos.DGV18.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_18()
        Frm_Expedientes_Activos.DGV18.Columns(0).HeaderText = "Id"
        Frm_Expedientes_Activos.DGV18.Columns(0).Width = 30
        Frm_Expedientes_Activos.DGV18.Columns(0).Visible = True
        Frm_Expedientes_Activos.DGV18.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV18.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Frm_Expedientes_Activos.DGV18.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Frm_Expedientes_Activos.DGV18.Columns(1).HeaderText = "ID_M_P"
        Frm_Expedientes_Activos.DGV18.Columns(1).Width = 0
        Frm_Expedientes_Activos.DGV18.Columns(1).Visible = False

        Frm_Expedientes_Activos.DGV18.Columns(2).HeaderText = "CODIGO"
        Frm_Expedientes_Activos.DGV18.Columns(2).Width = 0
        Frm_Expedientes_Activos.DGV18.Columns(2).Visible = False

        Frm_Expedientes_Activos.DGV18.Columns(3).HeaderText = "FECHA_ING_PAME"
        Frm_Expedientes_Activos.DGV18.Columns(3).Width = 0
        Frm_Expedientes_Activos.DGV18.Columns(3).Visible = False

        Frm_Expedientes_Activos.DGV18.Columns(4).HeaderText = "ID_CONT"
        Frm_Expedientes_Activos.DGV18.Columns(4).Width = 0
        Frm_Expedientes_Activos.DGV18.Columns(4).Visible = False

        Frm_Expedientes_Activos.DGV18.Columns(5).HeaderText = "Tipo"
        Frm_Expedientes_Activos.DGV18.Columns(5).Width = 200
        Frm_Expedientes_Activos.DGV18.Columns(5).Visible = True
        Frm_Expedientes_Activos.DGV18.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV18.Columns(6).HeaderText = "Desde"
        Frm_Expedientes_Activos.DGV18.Columns(6).Width = 80
        Frm_Expedientes_Activos.DGV18.Columns(6).Visible = True
        Frm_Expedientes_Activos.DGV18.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV18.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV18.Columns(7).HeaderText = "Hasta"
        Frm_Expedientes_Activos.DGV18.Columns(7).Width = 80
        Frm_Expedientes_Activos.DGV18.Columns(7).Visible = True
        Frm_Expedientes_Activos.DGV18.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV18.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV18.Columns(8).HeaderText = "Fir mado"
        Frm_Expedientes_Activos.DGV18.Columns(8).Width = 50
        Frm_Expedientes_Activos.DGV18.Columns(8).Visible = True
        Frm_Expedientes_Activos.DGV18.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV18.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV18.Columns(9).HeaderText = "Fecha Evaluacion"
        Frm_Expedientes_Activos.DGV18.Columns(9).Width = 80
        Frm_Expedientes_Activos.DGV18.Columns(9).Visible = True
        Frm_Expedientes_Activos.DGV18.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV18.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV18.Columns(10).HeaderText = "Calificacion Evaluacion"
        Frm_Expedientes_Activos.DGV18.Columns(10).Width = 80
        Frm_Expedientes_Activos.DGV18.Columns(10).Visible = True
        Frm_Expedientes_Activos.DGV18.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV18.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        'Frm_Expedientes_Activos.DGV18.Columns(10).DefaultCellStyle.BackColor = Color.Khaki

        Frm_Expedientes_Activos.DGV18.Columns(11).HeaderText = ""
        Frm_Expedientes_Activos.DGV18.Columns(11).Width = 0
        Frm_Expedientes_Activos.DGV18.Columns(11).Visible = False

        Frm_Expedientes_Activos.DGV18.Columns(12).HeaderText = "Clasificacion"
        Frm_Expedientes_Activos.DGV18.Columns(12).Width = 100
        Frm_Expedientes_Activos.DGV18.Columns(12).Visible = True
        Frm_Expedientes_Activos.DGV18.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV18.Columns(13).HeaderText = "Observaciones"
        Frm_Expedientes_Activos.DGV18.Columns(13).Width = 410
        Frm_Expedientes_Activos.DGV18.Columns(13).Visible = True
        Frm_Expedientes_Activos.DGV18.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_18()
        Dim I As Integer
        For I = 0 To Frm_Expedientes_Activos.DGV18.RowCount - 1
            If Frm_Expedientes_Activos.DGV18.Item(0, I).Value = mcyeID_M_CYE Then
                Frm_Expedientes_Activos.DGV18.Rows(I).Selected = True
                Frm_Expedientes_Activos.DGV18.CurrentCell = Frm_Expedientes_Activos.DGV18.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            If Me.TextBox3.Text = "" Then
                Me.TextBox3.Text = "0"
            End If
            If Not IsNumeric(Me.TextBox3.Text) Then
                Me.TextBox3.Text = "0"
            End If
            If IsNumeric(CInt(Me.TextBox3.Text)) Then
                Dim vNOTA As Double
                vNOTA = CDbl(Me.TextBox3.Text)
                If vNOTA < 60 Then
                    Me.ComboBox2.Text = "DEFICIENTE"
                End If
                If vNOTA > 59 And vNOTA < 69 Then
                    Me.ComboBox2.Text = "REGULAR"
                End If
                If vNOTA > 68 And vNOTA < 79 Then
                    Me.ComboBox2.Text = "BUENO"
                End If
                If vNOTA > 78 And vNOTA < 89 Then
                    Me.ComboBox2.Text = "MUY BUENO"
                End If
                If vNOTA > 88 And vNOTA < 101 Then
                    Me.ComboBox2.Text = "EXCELENTE"
                End If
            Else
                MsgBox("El dato ingresado no es númerico o no es váildo", vbInformation, "Mensaje del Sistema")
                Me.TextBox3.SelectAll()
                Me.TextBox3.Focus()
                Exit Sub
            End If
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
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
End Class