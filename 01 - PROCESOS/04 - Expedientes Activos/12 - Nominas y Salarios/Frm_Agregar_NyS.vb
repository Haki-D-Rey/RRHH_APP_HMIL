Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Agregar_NyS
    Private Sub Frm_Agregar_NyS_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = ""
            Me.DateTimePicker1.Value = Now
            Me.TextBox2.Text = "0"
            Me.TextBox3.Text = "0"
            Me.CheckBox1.Checked = False
            Me.TextBox4.Text = ""
            Me.ComboBox1.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Agregar_NyS_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = ""
        Call CARGAR_COMBO_CAT_DE_MOVIMIENTO_EN_NOMINA()
        Me.DateTimePicker1.Value = Now
        Call CARGAR_COMBO_CAT_DE_NOMINA()
        Me.ComboBox2.Text = "PAME"
        Me.TextBox2.Text = "0"
        Me.TextBox3.Text = "0"
        Me.CheckBox1.Checked = False
        Me.CheckBox2.Checked = False
        Me.TextBox4.Text = ""
        Me.ComboBox1.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox1.Text = ""
        Me.DateTimePicker1.Value = Now
        Me.TextBox2.Text = "0"
        Me.TextBox3.Text = "0"
        Me.CheckBox1.Checked = False
        Me.TextBox4.Text = ""
        Me.ComboBox1.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_MOVIMIENTO_EN_NOMINA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE MOVIMIENTO EN NOMINA] order by ID_MOVIMIENTO_N", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_MOVIMIENTO_N"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_NOMINA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE NOMINA] order by ID_NOMINA", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox2
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_NOMINA"
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
    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
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
    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
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
    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
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
            MiDataAdapter = New SqlDataAdapter("SELECT ID_M_NYS FROM [MAESTRO DE NOMINAS Y SALARIOS] ORDER BY ID_M_NYS", Conexion.CxRRHH)
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
            Me.Button2.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Tipo de Movimiento Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox2.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar la Nómina Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox2.Text) = 0 Or Me.TextBox2.Text = "" Or Me.TextBox2.Text = "0" Then
            MsgBox("Debe digitar el monto de la Propuesta Salarial del Empleado", vbInformation, "Mensaje del Sistema")
            Me.TextBox2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox3.Text) = 0 Or Me.TextBox3.Text = "" Or Me.TextBox3.Text = "0" Then
            MsgBox("Debe digitar el monto del Salario del Empleado", vbInformation, "Mensaje del Sistema")
            Me.TextBox3.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox4.Text) = 0 Then
            Me.TextBox4.Text = ""
        End If
        If Me.CheckBox2.Checked = True Then
            Call DESACTIVAR_SALARIOS()
        End If
        Call GRABAR_REGISTRO()
        vmnsID_M_NYS = Me.TextBox1.Text
        CADENAsql = "SELECT * FROM [VISTA MAESTRO DE NOMINAS Y SALARIOS] WHERE ID_M_P = " & Frm_Expedientes_Activos.TextBox35.Text & " ORDER BY ID_NOMINA, F_MOVIMIENTO_N DESC"
        Call CARGAR_DATAGRIDVIEW_11()
        Call ARMAR_DATAGRIDVIEW_11()
        Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_11()
        Me.TextBox1.Text = ""
        Me.DateTimePicker1.Value = Now
        Me.TextBox2.Text = "0"
        Me.TextBox3.Text = "0"
        Me.CheckBox1.Checked = False
        Me.TextBox4.Text = ""
        Me.ComboBox1.Focus()
    End Sub
    Private Sub DESACTIVAR_SALARIOS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE NOMINAS Y SALARIOS] SET SAL_ACTUAL = 'FALSE' WHERE ID_M_P = " & Frm_Expedientes_Activos.TextBox35.Text & " AND ID_NOMINA = " & Me.ComboBox2.SelectedValue & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub GRABAR_REGISTRO()
        Dim fFEC_MOV As String = Me.DateTimePicker1.Text
        Dim fechaMovimiento = "Convert(VARCHAR(10), Convert(Date, '" + fFEC_MOV + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE NOMINAS Y SALARIOS] (ID_M_NYS, ID_M_P, ID_MOVIMIENTO_N, F_MOVIMIENTO_N, ID_NOMINA, PROPUESTA, SALARIO, SAL_MINIMO, OBSERVACIONES, SAL_ACTUAL)" &
                                  "values (" & CInt(Me.TextBox1.Text) & ", " & Frm_Expedientes_Activos.TextBox35.Text & ", " & Me.ComboBox1.SelectedValue & ", " & fechaMovimiento & ", " & Me.ComboBox2.SelectedValue & ", '" & Me.TextBox2.Text & "', '" & Me.TextBox3.Text & "', '" & Me.CheckBox1.Checked & "', '" & Me.TextBox4.Text & "', '" & Me.CheckBox2.Checked & "')"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_11()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE NOMINAS Y SALARIOS]")
        Frm_Expedientes_Activos.DGV11.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE NOMINAS Y SALARIOS]")
        Frm_Expedientes_Activos.Label89.Text = Frm_Expedientes_Activos.DGV11.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_11()
        Frm_Expedientes_Activos.DGV11.Columns(0).HeaderText = "Id"
        Frm_Expedientes_Activos.DGV11.Columns(0).Width = 30
        Frm_Expedientes_Activos.DGV11.Columns(0).Visible = True
        Frm_Expedientes_Activos.DGV11.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV11.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Frm_Expedientes_Activos.DGV11.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Frm_Expedientes_Activos.DGV11.Columns(1).HeaderText = "ID_M_P"
        Frm_Expedientes_Activos.DGV11.Columns(1).Width = 0
        Frm_Expedientes_Activos.DGV11.Columns(1).Visible = False

        Frm_Expedientes_Activos.DGV11.Columns(2).HeaderText = "ID_MOVIMIENTO_N"
        Frm_Expedientes_Activos.DGV11.Columns(2).Width = 0
        Frm_Expedientes_Activos.DGV11.Columns(2).Visible = False

        Frm_Expedientes_Activos.DGV11.Columns(3).HeaderText = "Movimiento Nómina"
        Frm_Expedientes_Activos.DGV11.Columns(3).Width = 150
        Frm_Expedientes_Activos.DGV11.Columns(3).Visible = True
        Frm_Expedientes_Activos.DGV11.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV11.Columns(4).HeaderText = "Fecha"
        Frm_Expedientes_Activos.DGV11.Columns(4).Width = 70
        Frm_Expedientes_Activos.DGV11.Columns(4).Visible = True
        Frm_Expedientes_Activos.DGV11.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV11.Columns(5).HeaderText = "ID_NOMINA"
        Frm_Expedientes_Activos.DGV11.Columns(5).Width = 0
        Frm_Expedientes_Activos.DGV11.Columns(5).Visible = False

        Frm_Expedientes_Activos.DGV11.Columns(6).HeaderText = "Nomina"
        Frm_Expedientes_Activos.DGV11.Columns(6).Width = 210
        Frm_Expedientes_Activos.DGV11.Columns(6).Visible = True
        Frm_Expedientes_Activos.DGV11.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV11.Columns(7).HeaderText = "Propuesta"
        Frm_Expedientes_Activos.DGV11.Columns(7).Width = 70
        Frm_Expedientes_Activos.DGV11.Columns(7).Visible = True
        Frm_Expedientes_Activos.DGV11.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV11.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Frm_Expedientes_Activos.DGV11.Columns(7).DefaultCellStyle.Format = "##,##0.00"

        Frm_Expedientes_Activos.DGV11.Columns(8).HeaderText = "Salario"
        Frm_Expedientes_Activos.DGV11.Columns(8).Width = 70
        Frm_Expedientes_Activos.DGV11.Columns(8).Visible = True
        Frm_Expedientes_Activos.DGV11.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV11.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Frm_Expedientes_Activos.DGV11.Columns(8).DefaultCellStyle.Format = "##,##0.00"

        Frm_Expedientes_Activos.DGV11.Columns(9).HeaderText = "Salario Minimo"
        Frm_Expedientes_Activos.DGV11.Columns(9).Width = 60
        Frm_Expedientes_Activos.DGV11.Columns(9).Visible = True
        Frm_Expedientes_Activos.DGV11.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV11.Columns(10).HeaderText = "Observaciones"
        Frm_Expedientes_Activos.DGV11.Columns(10).Width = 400
        Frm_Expedientes_Activos.DGV11.Columns(10).Visible = True
        Frm_Expedientes_Activos.DGV11.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV11.Columns(11).HeaderText = "Activo"
        Frm_Expedientes_Activos.DGV11.Columns(11).Width = 50
        Frm_Expedientes_Activos.DGV11.Columns(11).Visible = True
        Frm_Expedientes_Activos.DGV11.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_11()
        Dim I As Integer
        For I = 0 To Frm_Expedientes_Activos.DGV11.RowCount - 1
            If Frm_Expedientes_Activos.DGV11.Item(0, I).Value = vmnsID_M_NYS Then
                Frm_Expedientes_Activos.DGV11.Rows(I).Selected = True
                Frm_Expedientes_Activos.DGV11.CurrentCell = Frm_Expedientes_Activos.DGV11.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub CheckBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
End Class