Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Aspirantes_Add_Procesos
    Private Sub Frm_Aspirantes_Add_Procesos_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.DateTimePicker1.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Aspirantes_Add_Procesos_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        RemoveHandler ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
        Call Clases.BUSCAR_FECHA_Y_HORA()
        Me.DateTimePicker1.Value = FECHA_SERVIDOR
        Call CARGAR_COMBO_SOL_CAT_DE_PROCESOS_PRINCIPALES()
        Call CARGAR_COMBO_SOL_CAT_DE_UBICACION_DE_RECURSOS()
        Me.CheckBox1.Checked = False
        Call CARGAR_COMBO_SOL_CAT_DE_RESULTADOS()
        Me.TextBox2.Text = ""
        Me.TextBox1.Text = ""
        Me.DateTimePicker1.Focus()
        AddHandler ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.DateTimePicker1.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_SOL_CAT_DE_PROCESOS_PRINCIPALES()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [SOL_CAT DE PROCESOS PRINCIPALES] ORDER BY ID_PROC", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_PROC"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_SOL_CAT_DE_UBICACION_DE_RECURSOS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [SOL_CAT DE UBICACION DE RECURSOS] ORDER BY ID_U_RECURSO", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox2
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_U_RECURSO"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_SOL_CAT_DE_RESULTADOS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [SOL_CAT DE RESULTADOS] WHERE ID_PROC = " & Me.ComboBox1.SelectedValue & " ORDER BY ID_R_SOL", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox3
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_R_SOL"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Call CARGAR_COMBO_SOL_CAT_DE_RESULTADOS()
    End Sub
    Private Sub DateTimePicker1_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker1.KeyDown
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
    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
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
    Private Sub ComboBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox3.KeyDown
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
        Dim CODIGO As Integer
        CODIGO = 0
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            MiDataAdapter = New SqlDataAdapter("SELECT ID_S_P FROM [SOL_MAESTRO DE SOLICITANTES PROCESOS] ORDER BY ID_S_P", Conexion.CxRRHH)
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
            MsgBox("Debe seleccionar el Proceso actual en el que se ecuentra el recurso", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox2.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar la Ubicación actual en la que se ecuentra el recurso", vbInformation, "Mensaje del Sistema")
            Me.ComboBox2.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox3.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Resultado del Proceso en el que se ecuentra el recurso", vbInformation, "Mensaje del Sistema")
            Me.ComboBox3.Focus()
            Exit Sub
        End If
        Call Me.GRABAR_REGISTRO()
        asppID_S_P = Me.TextBox1.Text
        If Me.ComboBox1.Text = "RECEPCION DE DOCUMENTOS" Then
            Call EDITAR_PRINCIPAL()
        End If
        Me.DateTimePicker1.Focus()
        CADENAsql = "SELECT * FROM [VISTA SOL_MAESTRO DE SOLICITANTES PROCESOS] WHERE ID_S = " & aspID_S & " ORDER BY FECHA_P"
        Call CARGAR_DATAGRIDVIEW_2()
        Call ARMAR_DATAGRIDVIEW_2()
        Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_2()
        CERRAR = False
    End Sub
    Private Sub EDITAR_PRINCIPAL()
        Dim fFEC_ING As String = Mid(Me.DateTimePicker1.Value, 1, 10)
        If Me.DateTimePicker1.Checked = True Then
            fFEC_ING = Mid(Me.DateTimePicker1.Value, 1, 10)
        Else
            fFEC_ING = "NULL"
        End If
        Dim fechaingreso As Object = If(fFEC_ING <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFEC_ING + "', 105), 23)", "NULL")

        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [SOL_MAESTRO DE SOLICITANTES] SET FECHA_I = " & fechaingreso & " WHERE ID_S = " & aspID_S & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_2()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA SOL_MAESTRO DE SOLICITANTES PROCESOS]")
        Frm_Aspirantes.DGV2.DataSource = MiDataSet.Tables("[VISTA SOL_MAESTRO DE SOLICITANTES PROCESOS]")
        'Me.Label2.Text = DGV1.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_2()
        Frm_Aspirantes.DGV2.Columns(0).HeaderText = "Id"
        Frm_Aspirantes.DGV2.Columns(0).Width = 10
        Frm_Aspirantes.DGV2.Columns(0).Visible = True
        Frm_Aspirantes.DGV2.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Aspirantes.DGV2.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Frm_Aspirantes.DGV2.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Frm_Aspirantes.DGV2.Columns(1).HeaderText = "ID_S"
        Frm_Aspirantes.DGV2.Columns(1).Width = 0
        Frm_Aspirantes.DGV2.Columns(1).Visible = False

        Frm_Aspirantes.DGV2.Columns(2).HeaderText = "Fecha Proceso"
        Frm_Aspirantes.DGV2.Columns(2).Width = 70
        Frm_Aspirantes.DGV2.Columns(2).Visible = True
        Frm_Aspirantes.DGV2.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Aspirantes.DGV2.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Aspirantes.DGV2.Columns(3).HeaderText = "ID_PROC"
        Frm_Aspirantes.DGV2.Columns(3).Width = 0
        Frm_Aspirantes.DGV2.Columns(3).Visible = False

        Frm_Aspirantes.DGV2.Columns(4).HeaderText = "Proceso"
        Frm_Aspirantes.DGV2.Columns(4).Width = 120
        Frm_Aspirantes.DGV2.Columns(4).Visible = True
        Frm_Aspirantes.DGV2.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Aspirantes.DGV2.Columns(4).Frozen = True

        Frm_Aspirantes.DGV2.Columns(5).HeaderText = "ID_U_RECURSO"
        Frm_Aspirantes.DGV2.Columns(5).Width = 0
        Frm_Aspirantes.DGV2.Columns(5).Visible = False

        Frm_Aspirantes.DGV2.Columns(6).HeaderText = "Ubicacion Recurso"
        Frm_Aspirantes.DGV2.Columns(6).Width = 80
        Frm_Aspirantes.DGV2.Columns(6).Visible = True
        Frm_Aspirantes.DGV2.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Aspirantes.DGV2.Columns(7).HeaderText = "Reali zado"
        Frm_Aspirantes.DGV2.Columns(7).Width = 40
        Frm_Aspirantes.DGV2.Columns(7).Visible = True
        Frm_Aspirantes.DGV2.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Aspirantes.DGV2.Columns(8).HeaderText = "ID_R_SOL"
        Frm_Aspirantes.DGV2.Columns(8).Width = 0
        Frm_Aspirantes.DGV2.Columns(8).Visible = False

        Frm_Aspirantes.DGV2.Columns(9).HeaderText = "Resultado"
        Frm_Aspirantes.DGV2.Columns(9).Width = 200
        Frm_Aspirantes.DGV2.Columns(9).Visible = True
        Frm_Aspirantes.DGV2.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Aspirantes.DGV2.Columns(10).HeaderText = "Observaciones"
        Frm_Aspirantes.DGV2.Columns(10).Width = 540
        Frm_Aspirantes.DGV2.Columns(10).Visible = True
        Frm_Aspirantes.DGV2.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Aspirantes.DGV2.Columns(11).HeaderText = "ID_ESTADO_R"
        Frm_Aspirantes.DGV2.Columns(11).Width = 0
        Frm_Aspirantes.DGV2.Columns(11).Visible = False
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_2()
        Dim I As Integer
        For I = 0 To Frm_Aspirantes.DGV2.RowCount - 1
            If Frm_Aspirantes.DGV2.Item(0, I).Value = asppID_S_P Then
                Frm_Aspirantes.DGV2.Rows(I).Selected = True
                Frm_Aspirantes.DGV2.CurrentCell = Frm_Aspirantes.DGV2.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub GRABAR_REGISTRO()
        Dim FECHA01 As String = Mid(Me.DateTimePicker1.Value, 1, 10)
        Dim f1 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA01 + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [SOL_MAESTRO DE SOLICITANTES PROCESOS] (ID_S_P, ID_S, FECHA_P, ID_PROC, ID_U_RECURSO, REALIZADA, ID_R_SOL, OBSERVACIONES, ID_ESTADO_R)" &
                                  "values (" & CInt(Me.TextBox1.Text) & ", " & aspID_S & ", " & f1 & ", " & Me.ComboBox1.SelectedValue & ", " & Me.ComboBox2.SelectedValue & ", '" & Me.CheckBox1.Checked & "', " & Me.ComboBox3.SelectedValue & ", '" & Me.TextBox2.Text & "', 1)"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
End Class