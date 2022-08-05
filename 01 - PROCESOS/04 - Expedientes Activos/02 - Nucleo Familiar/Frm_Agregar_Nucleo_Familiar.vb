Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Agregar_Nucleo_Familiar
    Private Sub Frm_Agregar_Nucleo_Familiar_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = ""
            Me.CheckBox5.Checked = False
            Me.TextBox2.Text = ""
            Me.DateTimePicker1.Value = Now
            Me.TextBox3.Text = ""
            Me.TextBox4.Text = ""
            Me.TextBox5.Text = ""
            Me.TextBox6.Text = ""
            Me.TextBox7.Text = ""
            Me.TextBox8.Text = ""
            Me.ComboBox1.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Agregar_Nucleo_Familiar_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = ""
        Call CARGAR_COMBO_CAT_DE_PARENTELA()
        Me.CheckBox5.Checked = False
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_SEXO()
        Me.TextBox2.Text = ""
        Me.DateTimePicker1.Value = Now
        Call CARGAR_COMBO_CAT_DE_NIVEL_ACADEMICO()
        Me.TextBox3.Text = Trim(Frm_Expedientes_Activos.TextBox7.Text)
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Me.TextBox6.Text = ""
        Me.TextBox7.Text = ""
        Me.TextBox8.Text = ""
        Me.ComboBox1.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox1.Text = ""
        Me.CheckBox5.Checked = False
        Me.TextBox2.Text = ""
        Me.DateTimePicker1.Value = Now
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Me.TextBox6.Text = ""
        Me.TextBox7.Text = ""
        Me.TextBox8.Text = ""
        Me.ComboBox1.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_PARENTELA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE PARENTELA] order by ID_PARENTELA", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_PARENTELA"
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
            With Me.ComboBox2
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
    Private Sub CARGAR_COMBO_CAT_DE_NIVEL_ACADEMICO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE NIVEL ACADEMICO] order by ID_N_ACADEMICO", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox3
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
    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
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
    Private Sub ComboBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox3.KeyDown
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
    Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyDown
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
            MiDataAdapter = New SqlDataAdapter("SELECT ID_NF FROM [MAESTRO DE NUCLEO FAMILIAR] ORDER BY ID_NF", Conexion.CxRRHH)
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
            MsgBox("Debe seleccionar el Tipo de Parentela Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox2.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Tipo de Sexo Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox2.Text) = 0 Then
            MsgBox("Debe digitar los Apellidos y Nombres de la Parentela", vbInformation, "Mensaje del Sistema")
            Me.TextBox2.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox3.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Nivel Academico Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox3.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox3.Text) = 0 Then
            MsgBox("Debe digitar la Dirección Domiciliar de la Parentela", vbInformation, "Mensaje del Sistema")
            Me.TextBox3.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox4.Text) = 0 Then
            Me.TextBox4.Text = ""
        End If
        If Len(Me.TextBox5.Text) = 0 Then
            Me.TextBox5.Text = ""
        End If
        If Len(Me.TextBox6.Text) = 0 Then
            Me.TextBox6.Text = ""
        End If
        If Len(Me.TextBox7.Text) = 0 Then
            Me.TextBox7.Text = ""
        End If
        If Len(Me.TextBox8.Text) = 0 Then
            Me.TextBox8.Text = ""
        End If
        Call GRABAR_REGISTRO()
        vmnfID_NF = Me.TextBox1.Text
        CADENAsql = "SELECT * FROM [VISTA MAESTRO DE NUCLEO FAMILIAR] WHERE ID_M_P = " & Frm_Expedientes_Activos.TextBox35.Text & " ORDER BY ID_M_P, ID_NF"
        Call CARGAR_DATAGRIDVIEW_1()
        Call ARMAR_DATAGRIDVIEW_1()
        Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_1()
        Me.TextBox2.Text = ""
        Me.TextBox6.Text = ""
        Me.TextBox7.Text = ""
        Me.TextBox8.Text = ""
        Me.TextBox1.Text = ""
        Me.ComboBox1.Focus()
        CERRAR = False
    End Sub
    Private Sub GRABAR_REGISTRO()
        Dim fFEC_NAC As String
        If Me.DateTimePicker1.Checked = True Then
            fFEC_NAC = Mid(Me.DateTimePicker1.Value, 1, 10)
        Else
            fFEC_NAC = "NULL"
        End If
        Dim fechaNacimiento As Object = If(fFEC_NAC <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFEC_NAC + "', 105), 23)", "NULL")

        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE NUCLEO FAMILIAR] (ID_NF, ID_M_P, ID_PARENTELA, APELLIDOS_NOMBRES, ID_T_SEXO, FEC_NAC, ID_N_ACADEMICO, DIRECCION_D, TELEFONO_1, TELEFONO_2, LUGAR_T, DIRECCION_T, CARGO_T, FALLECIDO)" &
                                  "values (" & CInt(Me.TextBox1.Text) & ", " & Frm_Expedientes_Activos.TextBox35.Text & ", " & Me.ComboBox1.SelectedValue & ", '" & Me.TextBox2.Text & "', " & Me.ComboBox2.SelectedValue & ", " & fechaNacimiento & ", " & Me.ComboBox3.SelectedValue & ", '" & Me.TextBox3.Text & "', '" & Me.TextBox4.Text & "', '" & Me.TextBox5.Text & "', '" & Me.TextBox6.Text & "', '" & Me.TextBox7.Text & "', '" & Me.TextBox8.Text & "', '" & Me.CheckBox5.Checked & "')"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_1()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE NUCLEO FAMILIAR]")
        Frm_Expedientes_Activos.DGV1.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE NUCLEO FAMILIAR]")
        Frm_Expedientes_Activos.Label61.Text = Frm_Expedientes_Activos.DGV1.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_1()
        Frm_Expedientes_Activos.DGV1.Columns(0).HeaderText = "Id"
        Frm_Expedientes_Activos.DGV1.Columns(0).Width = 30
        Frm_Expedientes_Activos.DGV1.Columns(0).Visible = True
        Frm_Expedientes_Activos.DGV1.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Frm_Expedientes_Activos.DGV1.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Frm_Expedientes_Activos.DGV1.Columns(1).HeaderText = "ID_M_P"
        Frm_Expedientes_Activos.DGV1.Columns(1).Width = 0
        Frm_Expedientes_Activos.DGV1.Columns(1).Visible = False

        Frm_Expedientes_Activos.DGV1.Columns(2).HeaderText = "ID_PARENTELA"
        Frm_Expedientes_Activos.DGV1.Columns(2).Width = 0
        Frm_Expedientes_Activos.DGV1.Columns(2).Visible = False

        Frm_Expedientes_Activos.DGV1.Columns(3).HeaderText = "Parentela"
        Frm_Expedientes_Activos.DGV1.Columns(3).Width = 60
        Frm_Expedientes_Activos.DGV1.Columns(3).Visible = True
        Frm_Expedientes_Activos.DGV1.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV1.Columns(4).HeaderText = "Apellidos y Nombres"
        Frm_Expedientes_Activos.DGV1.Columns(4).Width = 230
        Frm_Expedientes_Activos.DGV1.Columns(4).Visible = True
        Frm_Expedientes_Activos.DGV1.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV1.Columns(4).Frozen = True

        Frm_Expedientes_Activos.DGV1.Columns(5).HeaderText = "ID_T_SEXO"
        Frm_Expedientes_Activos.DGV1.Columns(5).Width = 0
        Frm_Expedientes_Activos.DGV1.Columns(5).Visible = False

        Frm_Expedientes_Activos.DGV1.Columns(6).HeaderText = "Sexo"
        Frm_Expedientes_Activos.DGV1.Columns(6).Width = 70
        Frm_Expedientes_Activos.DGV1.Columns(6).Visible = True
        Frm_Expedientes_Activos.DGV1.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV1.Columns(7).HeaderText = "Fecha Nacimiento"
        Frm_Expedientes_Activos.DGV1.Columns(7).Width = 70
        Frm_Expedientes_Activos.DGV1.Columns(7).Visible = True
        Frm_Expedientes_Activos.DGV1.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV1.Columns(8).HeaderText = "Edad"
        Frm_Expedientes_Activos.DGV1.Columns(8).Width = 30
        Frm_Expedientes_Activos.DGV1.Columns(8).Visible = True
        Frm_Expedientes_Activos.DGV1.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV1.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Frm_Expedientes_Activos.DGV1.Columns(9).HeaderText = "ID_N_ACADEMICO"
        Frm_Expedientes_Activos.DGV1.Columns(9).Width = 0
        Frm_Expedientes_Activos.DGV1.Columns(9).Visible = False

        Frm_Expedientes_Activos.DGV1.Columns(10).HeaderText = "Nivel Academico"
        Frm_Expedientes_Activos.DGV1.Columns(10).Width = 80
        Frm_Expedientes_Activos.DGV1.Columns(10).Visible = True
        Frm_Expedientes_Activos.DGV1.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV1.Columns(11).HeaderText = "Dirección Domiciliar"
        Frm_Expedientes_Activos.DGV1.Columns(11).Width = 250
        Frm_Expedientes_Activos.DGV1.Columns(11).Visible = True
        Frm_Expedientes_Activos.DGV1.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV1.Columns(12).HeaderText = "Telefono 1"
        Frm_Expedientes_Activos.DGV1.Columns(12).Width = 60
        Frm_Expedientes_Activos.DGV1.Columns(12).Visible = True
        Frm_Expedientes_Activos.DGV1.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV1.Columns(13).HeaderText = "Telefono 2"
        Frm_Expedientes_Activos.DGV1.Columns(13).Width = 60
        Frm_Expedientes_Activos.DGV1.Columns(13).Visible = True
        Frm_Expedientes_Activos.DGV1.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV1.Columns(14).HeaderText = "Lugar de Trabajo"
        Frm_Expedientes_Activos.DGV1.Columns(14).Width = 150
        Frm_Expedientes_Activos.DGV1.Columns(14).Visible = True
        Frm_Expedientes_Activos.DGV1.Columns(14).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV1.Columns(15).HeaderText = "Direccion de Trabajo"
        Frm_Expedientes_Activos.DGV1.Columns(15).Width = 150
        Frm_Expedientes_Activos.DGV1.Columns(15).Visible = True
        Frm_Expedientes_Activos.DGV1.Columns(15).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV1.Columns(16).HeaderText = "Cargo que Ocupa"
        Frm_Expedientes_Activos.DGV1.Columns(16).Width = 150
        Frm_Expedientes_Activos.DGV1.Columns(16).Visible = True
        Frm_Expedientes_Activos.DGV1.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV1.Columns(17).HeaderText = "Fallecido"
        Frm_Expedientes_Activos.DGV1.Columns(17).Width = 50
        Frm_Expedientes_Activos.DGV1.Columns(17).Visible = True
        Frm_Expedientes_Activos.DGV1.Columns(17).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_1()
        Dim I As Integer
        For I = 0 To Frm_Expedientes_Activos.DGV1.RowCount - 1
            If Frm_Expedientes_Activos.DGV1.Item(0, I).Value = vmnfID_NF Then
                Frm_Expedientes_Activos.DGV1.Rows(I).Selected = True
                Frm_Expedientes_Activos.DGV1.CurrentCell = Frm_Expedientes_Activos.DGV1.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
End Class