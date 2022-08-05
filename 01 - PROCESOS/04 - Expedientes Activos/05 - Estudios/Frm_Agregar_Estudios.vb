Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Agregar_Estudios
    Private Sub Frm_Agregar_Estudios_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = ""
            Me.TextBox2.Text = ""
            Me.DateTimePicker1.Value = Now
            Me.DateTimePicker1.Checked = False
            Me.TextBox3.Text = ""
            Me.TextBox4.Text = ""
            Me.TextBox5.Text = ""
            Me.CheckBox5.Checked = False
            Me.TextBox6.Text = ""
            Me.TextBox7.Text = ""
            Me.ComboBox1.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Agregar_Estudios_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        RemoveHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
        Me.TextBox1.Text = ""
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_ESTUDIO()
        Call CARGAR_COMBO_CAT_DE_ESPECIALIDADES_MEDICAS()
        Me.TextBox2.Text = ""
        Me.DateTimePicker1.Value = Now
        Me.DateTimePicker1.Checked = False
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Me.CheckBox5.Checked = False
        Call CARGAR_COMBO_CAT_DE_PAIS()
        Me.TextBox6.Text = ""
        Me.TextBox7.Text = ""
        Me.ComboBox1.Focus()
        AddHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.DateTimePicker1.Value = Now
        Me.DateTimePicker1.Checked = False
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Me.CheckBox5.Checked = False
        Me.TextBox6.Text = ""
        Me.TextBox7.Text = ""
        Me.ComboBox1.Focus()
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
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESPECIALIDADES MEDICAS] where ID_T_ESTUDIO = " & Me.ComboBox1.SelectedValue & " order by descripcion", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox3
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
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_ESTUDIO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE ESTUDIO] order by ID_T_ESTUDIO", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_T_ESTUDIO"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_PAIS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE PAIS] order by ID_PAIS", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox2
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_PAIS"
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
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
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
    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
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
    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
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
    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
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
    Private Sub GENERAR_CODIGO()
        CODIGO = 0
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            MiDataAdapter = New SqlDataAdapter("SELECT ID_M_EST FROM [MAESTRO DE ESTUDIOS] ORDER BY ID_M_EST", Conexion.CxRRHH)
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
            MsgBox("Debe seleccionar el Tipo de Estudio Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox3.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar la Especialidad Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox3.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox2.Text) = 0 Then
            MsgBox("Debe digitar el Ultimo Año Aprobado", vbInformation, "Mensaje del Sistema")
            Me.TextBox2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox3.Text) = 0 Then
            MsgBox("Debe digitar el Nombre del Estudio Corectamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox3.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox4.Text) = 0 Then
            MsgBox("Debe digitar el Nombre del Centro de Estudio Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox4.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox5.Text) = 0 Then
            MsgBox("Debe digitar la Dirección del Centro de Estudio Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox5.Focus()
            Exit Sub
        End If
        If Me.CheckBox5.Checked = True Then
            If Len(Me.ComboBox2.SelectedValue) = 0 Or Me.ComboBox2.Text = "." Then
                MsgBox("Debe seleccionar el Pais de Estudio Correctamente", vbInformation, "Mensaje del Sistema")
                Me.ComboBox2.Focus()
                Exit Sub
            End If
        End If
        If Me.CheckBox5.Checked = False Then
            If Me.ComboBox2.Text <> "SIN DEFINIR" Then
                Me.CheckBox5.Checked = True
                MsgBox("Se marcará la Ocpion <El Estudio es en el Extranjero?> debido a que se selecciono el Pais de Estudio.", vbInformation, "Mensaje del Sistema")
            End If
        End If
        If Len(Me.ComboBox2.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar correctamente el Pais de Estudio", vbInformation, "Mensaje del Sistema")
            Me.ComboBox2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox6.Text) = 0 Then
            MsgBox("Debe digitar el Nombre del Titulo Obtenido Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox6.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox7.Text) = 0 Then
            Me.TextBox7.Text = ""
        End If

        Call GRABAR_REGISTRO()
        vmesID_M_EST = Me.TextBox1.Text
        CADENAsql = "SELECT * FROM [VISTA MAESTRO DE ESTUDIOS] WHERE ID_M_P = " & Frm_Expedientes_Activos.TextBox35.Text & " ORDER BY ID_M_P, ID_M_EST"
        Call CARGAR_DATAGRIDVIEW_4()
        Call ARMAR_DATAGRIDVIEW_4()
        Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_4()
        CERRAR = False
    End Sub
    Private Sub GRABAR_REGISTRO()
        Dim fFEC_F_EST As String
        If Me.DateTimePicker1.Checked = True Then
            fFEC_F_EST = Mid(Me.DateTimePicker1.Value, 1, 10)
        Else
            fFEC_F_EST = "NULL"
        End If
        Dim fechaFEstudio As Object = If(fFEC_F_EST <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFEC_F_EST + "', 105), 23)", "NULL")

        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE ESTUDIOS] (ID_M_EST, ID_M_P, ID_T_ESTUDIO, ID_E_MED, ULT_GRADO_A, FECHA_ESTUDIO, NOMBRE_ESTUDIO, NOMBRE_C_ESTUDIO, DIRECC_C_ESTUDIO, ESTUDIO_F_PAIS, ID_PAIS, TITULO_OBT, OBSERVACIONES)" &
                                  "values (" & CInt(Me.TextBox1.Text) & ", " & Frm_Expedientes_Activos.TextBox35.Text & ", " & Me.ComboBox1.SelectedValue & ", " & Me.ComboBox3.SelectedValue & ", '" & Me.TextBox2.Text & "', " & fechaFEstudio & ", '" & Me.TextBox3.Text & "', '" & Me.TextBox4.Text & "', '" & Me.TextBox5.Text & "', '" & Me.CheckBox5.Checked & "', " & Me.ComboBox2.SelectedValue & ", '" & Me.TextBox6.Text & "', '" & Me.TextBox7.Text & "')"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_4()
        Dim I As Integer
        For I = 0 To Frm_Expedientes_Activos.DGV4.RowCount - 1
            If Frm_Expedientes_Activos.DGV4.Item(0, I).Value = vmesID_M_EST Then
                Frm_Expedientes_Activos.DGV4.Rows(I).Selected = True
                Frm_Expedientes_Activos.DGV4.CurrentCell = Frm_Expedientes_Activos.DGV4.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_4()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE ESTUDIOS]")
        Frm_Expedientes_Activos.DGV4.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE ESTUDIOS]")
        Frm_Expedientes_Activos.Label82.Text = Frm_Expedientes_Activos.DGV4.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_4()
        Frm_Expedientes_Activos.DGV4.Columns(0).HeaderText = "Id"
        Frm_Expedientes_Activos.DGV4.Columns(0).Width = 30
        Frm_Expedientes_Activos.DGV4.Columns(0).Visible = True
        Frm_Expedientes_Activos.DGV4.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV4.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Frm_Expedientes_Activos.DGV4.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Frm_Expedientes_Activos.DGV4.Columns(1).HeaderText = "ID_M_P"
        Frm_Expedientes_Activos.DGV4.Columns(1).Width = 0
        Frm_Expedientes_Activos.DGV4.Columns(1).Visible = False

        Frm_Expedientes_Activos.DGV4.Columns(2).HeaderText = "ID_T_ESTUDIO"
        Frm_Expedientes_Activos.DGV4.Columns(2).Width = 0
        Frm_Expedientes_Activos.DGV4.Columns(2).Visible = False

        Frm_Expedientes_Activos.DGV4.Columns(3).HeaderText = "Tipo de Estudio"
        Frm_Expedientes_Activos.DGV4.Columns(3).Width = 100
        Frm_Expedientes_Activos.DGV4.Columns(3).Visible = True
        Frm_Expedientes_Activos.DGV4.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV4.Columns(3).Frozen = True

        Frm_Expedientes_Activos.DGV4.Columns(4).HeaderText = "ID_E_MED"
        Frm_Expedientes_Activos.DGV4.Columns(4).Width = 0
        Frm_Expedientes_Activos.DGV4.Columns(4).Visible = False

        Frm_Expedientes_Activos.DGV4.Columns(5).HeaderText = "Especialidad"
        Frm_Expedientes_Activos.DGV4.Columns(5).Width = 150
        Frm_Expedientes_Activos.DGV4.Columns(5).Visible = True
        Frm_Expedientes_Activos.DGV4.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV4.Columns(6).HeaderText = "Ult. Grado"
        Frm_Expedientes_Activos.DGV4.Columns(6).Width = 50
        Frm_Expedientes_Activos.DGV4.Columns(6).Visible = True
        Frm_Expedientes_Activos.DGV4.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV4.Columns(7).HeaderText = "Fecha de Estudio"
        Frm_Expedientes_Activos.DGV4.Columns(7).Width = 80
        Frm_Expedientes_Activos.DGV4.Columns(7).Visible = True
        Frm_Expedientes_Activos.DGV4.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV4.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV4.Columns(8).HeaderText = "Nombre del Estudio"
        Frm_Expedientes_Activos.DGV4.Columns(8).Width = 280
        Frm_Expedientes_Activos.DGV4.Columns(8).Visible = True
        Frm_Expedientes_Activos.DGV4.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV4.Columns(9).HeaderText = "Nombre Centro del Estudio"
        Frm_Expedientes_Activos.DGV4.Columns(9).Width = 250
        Frm_Expedientes_Activos.DGV4.Columns(9).Visible = True
        Frm_Expedientes_Activos.DGV4.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV4.Columns(10).HeaderText = "Direccion Centro del Estudio"
        Frm_Expedientes_Activos.DGV4.Columns(10).Width = 200
        Frm_Expedientes_Activos.DGV4.Columns(10).Visible = True
        Frm_Expedientes_Activos.DGV4.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV4.Columns(11).HeaderText = "Estudio Extranjero"
        Frm_Expedientes_Activos.DGV4.Columns(11).Width = 70
        Frm_Expedientes_Activos.DGV4.Columns(11).Visible = True
        Frm_Expedientes_Activos.DGV4.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV4.Columns(12).HeaderText = "ID_PAIS"
        Frm_Expedientes_Activos.DGV4.Columns(12).Width = 0
        Frm_Expedientes_Activos.DGV4.Columns(12).Visible = False

        Frm_Expedientes_Activos.DGV4.Columns(13).HeaderText = "Pais"
        Frm_Expedientes_Activos.DGV4.Columns(13).Width = 100
        Frm_Expedientes_Activos.DGV4.Columns(13).Visible = True
        Frm_Expedientes_Activos.DGV4.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV4.Columns(14).HeaderText = "Titulo Obtenido"
        Frm_Expedientes_Activos.DGV4.Columns(14).Width = 280
        Frm_Expedientes_Activos.DGV4.Columns(14).Visible = True
        Frm_Expedientes_Activos.DGV4.Columns(14).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV4.Columns(15).HeaderText = "Observaciones"
        Frm_Expedientes_Activos.DGV4.Columns(15).Width = 280
        Frm_Expedientes_Activos.DGV4.Columns(15).Visible = True
        Frm_Expedientes_Activos.DGV4.Columns(15).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub ComboBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Call CARGAR_COMBO_CAT_DE_ESPECIALIDADES_MEDICAS()
    End Sub
End Class