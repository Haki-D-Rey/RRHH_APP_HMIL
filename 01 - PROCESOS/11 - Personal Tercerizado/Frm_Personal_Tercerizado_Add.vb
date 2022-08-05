Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Public Class Frm_Personal_Tercerizado_Add
    Private Sub Frm_Personal_Tercerizado_Add_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.ComboBox1.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Personal_Tercerizado_Add_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = "0"
        Call CARGAR_COMBO_CAT_DE_EMPRESAS_TERCERIZADAS()
        Me.ComboBox1.Text = Frm_Personal_Tercerizado.ComboBox1.Text
        Me.TextBox2.Text = ""
        Me.TextBox11.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_SEXO()
        Me.TextBox6.Text = ""
        Me.TextBox5.Text = ""
        Me.TextBox7.Text = ""
        Call CARGAR_COMBO_CAT_DE_ESTADO_DE_PERSONAS()
        Call Clases.BUSCAR_FECHA_Y_HORA()
        Me.DateTimePicker1.Value = Mid(FECHA_SERVIDOR, 1, 10)
        Me.DateTimePicker1.Checked = False
        Me.DateTimePicker2.Value = Mid(FECHA_SERVIDOR, 1, 10)
        Me.DateTimePicker2.Checked = False
        Me.TextBox10.Text = ""
        Me.TextBox14.Text = ""
        Me.ComboBox1.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.ComboBox1.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTADO_DE_PERSONAS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTADO DE PERSONAS] ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox4
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
    Private Sub CARGAR_COMBO_CAT_DE_EMPRESAS_TERCERIZADAS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE EMPRESAS TERCERIZADAS] ORDER BY ID_E_T", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_E_T"
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
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE SEXO] ORDER BY ID_T_SEXO", Conexion.CxRRHH)
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
    Private Sub TextBox11_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox11.KeyDown
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
    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox4.KeyDown
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
    Private Sub TextBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox9.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Dim CODIGO_ID As Integer
    Private Sub GENERAR_ID()
        CODIGO_ID = 900000000
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            MiDataAdapter = New SqlDataAdapter("SELECT ID_P_T FROM [MAESTRO DE PERSONAL TERCERIZADO] ORDER BY ID_P_T", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                For I = 0 To MiDataTable.Rows.Count - 1
                    CODIGO_ID = MiDataTable.Rows(I).Item(0).ToString + 1
                Next
            Else
                CODIGO_ID = 900000000 + 1
            End If
            Me.TextBox1.Text = CODIGO_ID
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim CODIGO As String
    Private Sub GENERAR_CODIGO()
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            MiDataAdapter = New SqlDataAdapter("SELECT CODIGO FROM [MAESTRO DE PERSONAL TERCERIZADO] ORDER BY CODIGO", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                For I = 0 To MiDataTable.Rows.Count - 1
                    CODIGO = MiDataTable.Rows(I).Item(0).ToString
                Next
            Else
                CODIGO = 900000001
            End If
            CODIGO = CInt(CODIGO) + 1
            Me.TextBox2.Text = Format(CInt(CODIGO), "0000000000")
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
        Call GENERAR_ID()
        If TextBox1.Text = "0" Or TextBox1.Text = "" Then
            MsgBox("No se ha generado el Codigo del registro", vbInformation, "Mensaje del Sistema")
            Me.TextBox1.Text = "0"
            Me.Button2.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar la Empresa Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        Me.TextBox2.Text = "0000000000"
        Call GENERAR_CODIGO()
        If TextBox2.Text = "0" Or TextBox2.Text = "" Then
            MsgBox("No se ha generado el Codigo de la Persona", vbInformation, "Mensaje del Sistema")
            Me.TextBox2.Text = "0"
            Me.Button2.Focus()
            Exit Sub
        End If
        If TextBox11.Text = "" Then
            MsgBox("Debe Digitar los Apellidos Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox11.Focus()
            Exit Sub
        End If
        If TextBox3.Text = "" Then
            MsgBox("Debe Digitar los Nombres Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox3.Focus()
            Exit Sub
        End If
        If TextBox4.Text = "" Then
            MsgBox("Debe Digitar la Cedula Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox4.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox2.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Tipo de Sexo Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox2.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox4.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Estado de la Persona Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox4.Focus()
            Exit Sub
        End If
        If TextBox14.Text = "" Then
            Me.TextBox14.Text = "."
        End If
        Call Me.GRABAR_REGISTRO()
        Frm_Personal_Tercerizado.ComboBox1.Text = Me.ComboBox1.Text
        CERRAR = False
        ptID_P_T = Me.TextBox1.Text
        ptCODIGO = Me.TextBox2.Text
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
        Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        Me.TextBox1.Text = "0"
        Me.TextBox2.Text = ""
        Me.TextBox11.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox6.Text = ""
        Me.TextBox5.Text = ""
        Me.TextBox7.Text = ""
        Me.TextBox8.Text = ""
        Me.TextBox9.Text = ""
        Me.TextBox14.Text = ""
        Me.PictureBox4.Image = Me.PictureBox7.Image
        Me.ComboBox1.Focus()
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter("Select ID_P_T, ID_E_T, D_E_T, CODIGO, APELLIDOS, NOMBRES, AN, N_CEDULA, ID_T_SEXO, D_T_SEXO, TELEFONO_1, TELEFONO_2, D_CARGO_ES, ID_ESTADO_P, D_ESTADO_P, DIRECCION, NSS, OBSERVACIONES from [VISTA MAESTRO DE PERSONAL TERCERIZADO] WHERE ID_E_T = " & Me.ComboBox1.SelectedValue & " ORDER BY CODIGO", Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE PERSONAL TERCERIZADO]")
        Frm_Personal_Tercerizado.DGV.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE PERSONAL TERCERIZADO]")
        Frm_Personal_Tercerizado.Label2.Text = Frm_Personal_Tercerizado.DGV.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW()
        Frm_Personal_Tercerizado.DGV.Columns(0).HeaderText = "Id"
        Frm_Personal_Tercerizado.DGV.Columns(0).Width = 30
        Frm_Personal_Tercerizado.DGV.Columns(0).Visible = True
        Frm_Personal_Tercerizado.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Personal_Tercerizado.DGV.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Frm_Personal_Tercerizado.DGV.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Frm_Personal_Tercerizado.DGV.Columns(1).HeaderText = "ID_E_T"
        Frm_Personal_Tercerizado.DGV.Columns(1).Width = 0
        Frm_Personal_Tercerizado.DGV.Columns(1).Visible = False

        Frm_Personal_Tercerizado.DGV.Columns(2).HeaderText = "Empresa"
        Frm_Personal_Tercerizado.DGV.Columns(2).Width = 0
        Frm_Personal_Tercerizado.DGV.Columns(2).Visible = False
        Frm_Personal_Tercerizado.DGV.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Personal_Tercerizado.DGV.Columns(3).HeaderText = "Codigo"
        Frm_Personal_Tercerizado.DGV.Columns(3).Width = 80
        Frm_Personal_Tercerizado.DGV.Columns(3).Visible = True
        Frm_Personal_Tercerizado.DGV.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Personal_Tercerizado.DGV.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Personal_Tercerizado.DGV.Columns(4).HeaderText = "APELLIDOS"
        Frm_Personal_Tercerizado.DGV.Columns(4).Width = 0
        Frm_Personal_Tercerizado.DGV.Columns(4).Visible = False

        Frm_Personal_Tercerizado.DGV.Columns(5).HeaderText = "NOMBRES"
        Frm_Personal_Tercerizado.DGV.Columns(5).Width = 0
        Frm_Personal_Tercerizado.DGV.Columns(5).Visible = False

        Frm_Personal_Tercerizado.DGV.Columns(6).HeaderText = "Apellidos y Nombres"
        Frm_Personal_Tercerizado.DGV.Columns(6).Width = 230
        Frm_Personal_Tercerizado.DGV.Columns(6).Visible = True
        Frm_Personal_Tercerizado.DGV.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Personal_Tercerizado.DGV.Columns(7).HeaderText = "Cedula"
        Frm_Personal_Tercerizado.DGV.Columns(7).Width = 110
        Frm_Personal_Tercerizado.DGV.Columns(7).Visible = True
        Frm_Personal_Tercerizado.DGV.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Personal_Tercerizado.DGV.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Personal_Tercerizado.DGV.Columns(8).HeaderText = "ID_T_SEXO"
        Frm_Personal_Tercerizado.DGV.Columns(8).Width = 0
        Frm_Personal_Tercerizado.DGV.Columns(8).Visible = False

        Frm_Personal_Tercerizado.DGV.Columns(9).HeaderText = "Sexo"
        Frm_Personal_Tercerizado.DGV.Columns(9).Width = 70
        Frm_Personal_Tercerizado.DGV.Columns(9).Visible = True
        Frm_Personal_Tercerizado.DGV.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Personal_Tercerizado.DGV.Columns(9).Frozen = True

        Frm_Personal_Tercerizado.DGV.Columns(10).HeaderText = "Telefono (Claro)"
        Frm_Personal_Tercerizado.DGV.Columns(10).Width = 80
        Frm_Personal_Tercerizado.DGV.Columns(10).Visible = True
        Frm_Personal_Tercerizado.DGV.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Personal_Tercerizado.DGV.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Personal_Tercerizado.DGV.Columns(11).HeaderText = "Telefono (Tigo)"
        Frm_Personal_Tercerizado.DGV.Columns(11).Width = 80
        Frm_Personal_Tercerizado.DGV.Columns(11).Visible = True
        Frm_Personal_Tercerizado.DGV.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Personal_Tercerizado.DGV.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Personal_Tercerizado.DGV.Columns(12).HeaderText = "Cargo"
        Frm_Personal_Tercerizado.DGV.Columns(12).Width = 200
        Frm_Personal_Tercerizado.DGV.Columns(12).Visible = True
        Frm_Personal_Tercerizado.DGV.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Personal_Tercerizado.DGV.Columns(13).HeaderText = "ID_ESTADO_P"
        Frm_Personal_Tercerizado.DGV.Columns(13).Width = 0
        Frm_Personal_Tercerizado.DGV.Columns(13).Visible = False

        Frm_Personal_Tercerizado.DGV.Columns(14).HeaderText = "Estado"
        Frm_Personal_Tercerizado.DGV.Columns(14).Width = 70
        Frm_Personal_Tercerizado.DGV.Columns(14).Visible = True
        Frm_Personal_Tercerizado.DGV.Columns(14).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Personal_Tercerizado.DGV.Columns(15).HeaderText = "Direccion"
        Frm_Personal_Tercerizado.DGV.Columns(15).Width = 250
        Frm_Personal_Tercerizado.DGV.Columns(15).Visible = True
        Frm_Personal_Tercerizado.DGV.Columns(15).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Personal_Tercerizado.DGV.Columns(16).HeaderText = "Nss"
        Frm_Personal_Tercerizado.DGV.Columns(16).Width = 50
        Frm_Personal_Tercerizado.DGV.Columns(16).Visible = True
        Frm_Personal_Tercerizado.DGV.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Personal_Tercerizado.DGV.Columns(17).HeaderText = "Observaciones"
        Frm_Personal_Tercerizado.DGV.Columns(17).Width = 250
        Frm_Personal_Tercerizado.DGV.Columns(17).Visible = True
        Frm_Personal_Tercerizado.DGV.Columns(17).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Personal_Tercerizado.DGV.Columns(18).HeaderText = "Fecha Ingreso"
        Frm_Personal_Tercerizado.DGV.Columns(18).Width = 80
        Frm_Personal_Tercerizado.DGV.Columns(18).Visible = True
        Frm_Personal_Tercerizado.DGV.Columns(18).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Personal_Tercerizado.DGV.Columns(19).HeaderText = "Fecha Baja"
        Frm_Personal_Tercerizado.DGV.Columns(19).Width = 80
        Frm_Personal_Tercerizado.DGV.Columns(19).Visible = True
        Frm_Personal_Tercerizado.DGV.Columns(19).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Personal_Tercerizado.DGV.Columns(20).HeaderText = "Causa de Retiro"
        Frm_Personal_Tercerizado.DGV.Columns(20).Width = 150
        Frm_Personal_Tercerizado.DGV.Columns(20).Visible = True
        Frm_Personal_Tercerizado.DGV.Columns(20).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft

    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        Dim I As Integer
        For I = 0 To Frm_Personal_Tercerizado.DGV.RowCount - 1
            If Frm_Personal_Tercerizado.DGV.Item(3, I).Value = ptCODIGO Then
                Frm_Personal_Tercerizado.DGV.Rows(I).Selected = True
                Frm_Personal_Tercerizado.DGV.CurrentCell = Frm_Personal_Tercerizado.DGV.Item(3, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub GRABAR_REGISTRO()
        Dim fFEC_INGRESO As String
        If Me.DateTimePicker1.Checked = True Then
            fFEC_INGRESO = Mid(Me.DateTimePicker1.Value, 1, 10)
        Else
            fFEC_INGRESO = "NULL"
        End If
        Dim fechaingreso As Object = If(fFEC_INGRESO <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFEC_INGRESO + "', 105), 23)", "NULL")

        Dim fFEC_BAJA As String
        If Me.DateTimePicker2.Checked = True Then
            fFEC_BAJA = Mid(Me.DateTimePicker2.Value, 1, 10)
        Else
            fFEC_BAJA = "NULL"
        End If
        Dim fechabaja As Object = If(fFEC_BAJA <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFEC_BAJA + "', 105), 23)", "NULL")


        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim SQL As String
        Dim ms As New System.IO.MemoryStream() 'creando el memoryStream
        Try
            Me.PictureBox4.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
            SQL = "INSERT INTO [MAESTRO DE PERSONAL TERCERIZADO] (ID_P_T, ID_E_T, CODIGO, APELLIDOS, NOMBRES, N_CEDULA, ID_T_SEXO, TELEFONO_1, TELEFONO_2, D_CARGO_ES, ID_ESTADO_P, DIRECCION, NSS, FOTO, OBSERVACIONES, F_INGRESO, F_BAJA, CAUSA_RETIRO) values (" & CInt(Me.TextBox1.Text) & ", " & Me.ComboBox1.SelectedValue & ", '" & Trim(Me.TextBox2.Text) & "', '" & Trim(Me.TextBox11.Text) & "', '" & Trim(Me.TextBox3.Text) & "', '" & Trim(Me.TextBox4.Text) & "', " & Me.ComboBox2.SelectedValue & ", '" & Trim(Me.TextBox6.Text) & "', '" & Trim(Me.TextBox5.Text) & "', '" & Trim(Me.TextBox7.Text) & "', " & Me.ComboBox4.SelectedValue & ", '" & Trim(Me.TextBox9.Text) & "', '" & Trim(Me.TextBox8.Text) & "',  @data, '" & Trim(Me.TextBox14.Text) & "', " & fechaingreso & ", " & fechabaja & ", '" & Me.TextBox10.Text & "')"
            Dim cmd As New SqlCommand(SQL, CxRRHH)
            cmd.Parameters.Add(New SqlParameter("@data", ms.GetBuffer()))
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
    Dim CONTROL_FOTO As Boolean
    Dim data() As Byte 'arreglo de bytes el cual contedra la imagen
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim ios As FileStream 'Manejo de archivos
        Try
            Me.OpenFileDialog1.Filter = "Imagenes (JPG)|*.jpg" 'filtro de archivos del OpenFileDialog 
            If Me.OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.Cancel Then ' en caso de que se aplaste el boton cancelar salga y no haga nada
                CONTROL_FOTO = False
                Exit Sub
            Else
                ios = New FileStream(Me.OpenFileDialog1.FileName, FileMode.Open, FileAccess.Read) 'instanciamos en Stream indicandole la ruta del arvhivo,el modo de acceso y si es de lectura o escritura
                ReDim Data(ios.Length) 'llenamos el arreglo
                ios.Read(Data, 0, CInt(ios.Length)) 'llenamos el arreglo
                Me.PictureBox4.SizeMode = PictureBoxSizeMode.StretchImage 'establecemos como se visualiza la imagen
                Me.PictureBox4.Load(Me.OpenFileDialog1.FileName) 'visualizamos abriendo el archivo seleccionado
                CONTROL_FOTO = True
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Mensaje del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification) 'en caso de error muestre un mensaje
        End Try
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
    Private Sub TextBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox10.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
End Class