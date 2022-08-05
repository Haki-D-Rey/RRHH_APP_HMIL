Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Public Class Frm_Medicos_Acreditados_Agregar
    Private Sub Frm_Medicos_Acreditados_Agregar_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox2.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Medicos_Acreditados_Agregar_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Call INICIAR_OBJETOS()
        Me.TextBox6.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox2.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub INICIAR_OBJETOS()
        Me.TextBox1.Text = "0"
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Call CARGAR_COMBO_CAT_DE_ESPECIALIDADES_MEDICAS()
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Me.TextBox6.Text = ""
        Me.TextBox7.Text = ""
        Me.TextBox8.Text = ""
        Me.TextBox9.Text = ""
        Call CARGAR_COMBO_CAT_DE_NACIONALIDAD_D()
        Me.TextBox10.Text = ""
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_SEXO()
        Me.TextBox11.Text = ""
        Me.TextBox12.Text = ""
        Me.TextBox13.Text = ""
        Me.TextBox14.Text = ""
        Me.PictureBox4.Image = Nothing
        Me.TextBox2.Focus()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_SEXO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE SEXO] ORDER BY ID_T_SEXO", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox3
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
    Private Sub CARGAR_COMBO_CAT_DE_NACIONALIDAD_D()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE NACIONALIDAD_D] ORDER BY ID_NACIONALIDAD_D", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox2
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_NACIONALIDAD_D"
            End With
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESPECIALIDADES_MEDICAS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESPECIALIDADES MEDICAS] ORDER BY ID_E_MED", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
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
    Private Sub TextBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox9.KeyDown
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
    Private Sub TextBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox10.KeyDown
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
    Private Sub DateTimePicker3_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker3.KeyDown
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
    Private Sub TextBox14_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox14.KeyDown
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
            MiDataAdapter = New SqlDataAdapter("SELECT ID_MA FROM [MAESTRO DE MEDICOS ACREDITADOS] ORDER BY ID_MA", Conexion.CxRRHH)
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
        If TextBox2.Text = "" Then
            MsgBox("Debe Digitar los Apellidos Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox2.Focus()
            Exit Sub
        End If
        If TextBox3.Text = "" Then
            MsgBox("Debe Digitar los Nombres Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox3.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar la Especialidad Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If TextBox4.Text = "" Then
            MsgBox("Debe Digitar el Número de Cédula Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox4.Focus()
            Exit Sub
        End If
        If TextBox5.Text = "" Then
            MsgBox("Debe Digitar el Número MINSA Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox5.Focus()
            Exit Sub
        End If
        If TextBox6.Text = "" Then
            TextBox6.Text = ""
        End If
        If TextBox7.Text = "" Then
            TextBox7.Text = ""
        End If
        If Len(Me.ComboBox2.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar la Nacionalidad Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox2.Focus()
            Exit Sub
        End If
        If TextBox8.Text = "" Then
            TextBox8.Text = ""
        End If
        If TextBox9.Text = "" Then
            TextBox9.Text = ""
        End If
        If Len(Me.ComboBox3.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Tipo de Sexo Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox3.Focus()
            Exit Sub
        End If
        If TextBox10.Text = "" Then
            TextBox10.Text = ""
        End If
        If TextBox11.Text = "" Then
            TextBox11.Text = ""
        End If
        If TextBox12.Text = "" Then
            TextBox12.Text = ""
        End If
        If TextBox13.Text = "" Then
            TextBox13.Text = ""
        End If
        If TextBox14.Text = "" Then
            TextBox14.Text = ""
        End If
        Call Me.GRABAR_REGISTRO()
        CERRAR = False
        vMAc_ID_MA = Me.TextBox1.Text
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
        Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        Me.TextBox1.Text = "0"
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Me.TextBox6.Text = ""
        Me.TextBox7.Text = ""
        Me.TextBox8.Text = ""
        Me.TextBox9.Text = ""
        Me.TextBox10.Text = ""
        Me.TextBox11.Text = ""
        Me.TextBox12.Text = ""
        Me.TextBox13.Text = ""
        Me.TextBox14.Text = ""
        Me.TextBox2.Focus()
    End Sub
    Dim fFEC_1 As String
    Dim fFEC_2 As String
    Dim fFEC_3 As String
    Private Sub GRABAR_REGISTRO()
        If Me.DateTimePicker1.Checked = True Then
            fFEC_1 = Mid(Me.DateTimePicker1.Value, 1, 10)
        Else
            fFEC_1 = "NULL"
        End If
        Dim F1 As Object = If(fFEC_1 <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFEC_1 + "', 105), 23)", "NULL")

        If Me.DateTimePicker2.Checked = True Then
            fFEC_2 = Mid(Me.DateTimePicker2.Value, 1, 10)
        Else
            fFEC_2 = "NULL"
        End If
        Dim F2 As Object = If(fFEC_2 <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFEC_2 + "', 105), 23)", "NULL")

        If Me.DateTimePicker3.Checked = True Then
            fFEC_3 = Mid(Me.DateTimePicker3.Value, 1, 10)
        Else
            fFEC_3 = "NULL"
        End If
        Dim F3 As Object = If(fFEC_3 <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFEC_3 + "', 105), 23)", "NULL")

        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()

        Dim SQL As String
        Dim ms As New System.IO.MemoryStream() 'creando el memoryStream
        Try
            If CONTROL_FOTO = True Then
                Me.PictureBox4.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
            Else
                Me.PictureBox7.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
            End If

            SQL = "INSERT INTO [MAESTRO DE MEDICOS ACREDITADOS] (ID_MA, APELLIDOS, NOMBRES,  ID_E_MED, N_CEDULA, N_MINSA, TELEFONO_1, TELEFONO_2, CORREO_1, CORREO_2, ID_NACIONALIDAD_D, DIRECCION_D, ID_T_SEXO, FEC_NAC, FEC_DESDE, FEC_HASTA, FACEBOOK, TWITTER, WHATSAPP, OBSERVACIONES, FOTO)" &
                  "values (" & CInt(Me.TextBox1.Text) & ",  '" & Me.TextBox2.Text & "', '" & Me.TextBox3.Text & "', " & Me.ComboBox1.SelectedValue & ",  '" & Me.TextBox4.Text & "', '" & Me.TextBox5.Text & "', '" & Me.TextBox6.Text & "', '" & Me.TextBox7.Text & "', '" & Me.TextBox8.Text & "', '" & Me.TextBox9.Text & "', " & Me.ComboBox2.SelectedValue & ", '" & Me.TextBox10.Text & "',  " & Me.ComboBox3.SelectedValue & ", " & F1 & ", " & F2 & ", " & F3 & ", '" & Me.TextBox11.Text & "',  '" & Me.TextBox12.Text & "',  '" & Me.TextBox13.Text & "',  '" & Me.TextBox14.Text & "', @data)"

            Dim cmd As New SqlCommand(SQL, CxRRHH)
            cmd.Parameters.Add(New SqlParameter("@data", ms.GetBuffer()))
            cmd.ExecuteNonQuery()
            cmd = Nothing
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter("Select * from [VISTA MAESTRO DE MEDICOS ACREDITADOS] ORDER BY APELLIDOS, NOMBRES", Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE MEDICOS ACREDITADOS]")
        Frm_Medicos_Acreditados.DGV.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE MEDICOS ACREDITADOS]")
        Frm_Medicos_Acreditados.Label4.Text = "Total de Registros: " & Frm_Medicos_Acreditados.DGV.RowCount
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW()
        Frm_Medicos_Acreditados.DGV.Columns(0).HeaderText = "Id"
        Frm_Medicos_Acreditados.DGV.Columns(0).Width = 30
        Frm_Medicos_Acreditados.DGV.Columns(0).Visible = True
        Frm_Medicos_Acreditados.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Medicos_Acreditados.DGV.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Frm_Medicos_Acreditados.DGV.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Frm_Medicos_Acreditados.DGV.Columns(1).HeaderText = "Apellidos"
        Frm_Medicos_Acreditados.DGV.Columns(1).Width = 150
        Frm_Medicos_Acreditados.DGV.Columns(1).Visible = True
        Frm_Medicos_Acreditados.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Medicos_Acreditados.DGV.Columns(2).HeaderText = "Nombres"
        Frm_Medicos_Acreditados.DGV.Columns(2).Width = 150
        Frm_Medicos_Acreditados.DGV.Columns(2).Visible = True
        Frm_Medicos_Acreditados.DGV.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Medicos_Acreditados.DGV.Columns(3).HeaderText = "ID_E_MED"
        Frm_Medicos_Acreditados.DGV.Columns(3).Width = 0
        Frm_Medicos_Acreditados.DGV.Columns(3).Visible = False

        Frm_Medicos_Acreditados.DGV.Columns(4).HeaderText = "Especialidad Medica"
        Frm_Medicos_Acreditados.DGV.Columns(4).Width = 170
        Frm_Medicos_Acreditados.DGV.Columns(4).Visible = True
        Frm_Medicos_Acreditados.DGV.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Medicos_Acreditados.DGV.Columns(4).Frozen = True

        Frm_Medicos_Acreditados.DGV.Columns(5).HeaderText = "Cedula"
        Frm_Medicos_Acreditados.DGV.Columns(5).Width = 100
        Frm_Medicos_Acreditados.DGV.Columns(5).Visible = True
        Frm_Medicos_Acreditados.DGV.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Medicos_Acreditados.DGV.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Medicos_Acreditados.DGV.Columns(5).DefaultCellStyle.BackColor = Color.Khaki

        Frm_Medicos_Acreditados.DGV.Columns(6).HeaderText = "No. MINSA"
        Frm_Medicos_Acreditados.DGV.Columns(6).Width = 50
        Frm_Medicos_Acreditados.DGV.Columns(6).Visible = True
        Frm_Medicos_Acreditados.DGV.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Medicos_Acreditados.DGV.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Medicos_Acreditados.DGV.Columns(6).DefaultCellStyle.BackColor = Color.Khaki

        Frm_Medicos_Acreditados.DGV.Columns(7).HeaderText = "Telefono 1"
        Frm_Medicos_Acreditados.DGV.Columns(7).Width = 80
        Frm_Medicos_Acreditados.DGV.Columns(7).Visible = True
        Frm_Medicos_Acreditados.DGV.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Medicos_Acreditados.DGV.Columns(8).HeaderText = "Telefono 2"
        Frm_Medicos_Acreditados.DGV.Columns(8).Width = 80
        Frm_Medicos_Acreditados.DGV.Columns(8).Visible = True
        Frm_Medicos_Acreditados.DGV.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Medicos_Acreditados.DGV.Columns(9).HeaderText = "CORREO_1"
        Frm_Medicos_Acreditados.DGV.Columns(9).Width = 0
        Frm_Medicos_Acreditados.DGV.Columns(9).Visible = False

        Frm_Medicos_Acreditados.DGV.Columns(10).HeaderText = "CORREO_2"
        Frm_Medicos_Acreditados.DGV.Columns(10).Width = 0
        Frm_Medicos_Acreditados.DGV.Columns(10).Visible = False

        Frm_Medicos_Acreditados.DGV.Columns(11).HeaderText = "ID_NACIONALIDAD_D"
        Frm_Medicos_Acreditados.DGV.Columns(11).Width = 0
        Frm_Medicos_Acreditados.DGV.Columns(11).Visible = False

        Frm_Medicos_Acreditados.DGV.Columns(12).HeaderText = "Nacionalidad"
        Frm_Medicos_Acreditados.DGV.Columns(12).Width = 100
        Frm_Medicos_Acreditados.DGV.Columns(12).Visible = True
        Frm_Medicos_Acreditados.DGV.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Medicos_Acreditados.DGV.Columns(13).HeaderText = "Direccion"
        Frm_Medicos_Acreditados.DGV.Columns(13).Width = 170
        Frm_Medicos_Acreditados.DGV.Columns(13).Visible = True
        Frm_Medicos_Acreditados.DGV.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Medicos_Acreditados.DGV.Columns(14).HeaderText = "ID_T_SEXO"
        Frm_Medicos_Acreditados.DGV.Columns(14).Width = 0
        Frm_Medicos_Acreditados.DGV.Columns(14).Visible = False

        Frm_Medicos_Acreditados.DGV.Columns(15).HeaderText = "Sexo"
        Frm_Medicos_Acreditados.DGV.Columns(15).Width = 70
        Frm_Medicos_Acreditados.DGV.Columns(15).Visible = True
        Frm_Medicos_Acreditados.DGV.Columns(15).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Medicos_Acreditados.DGV.Columns(16).HeaderText = "Fecha Nacimiento"
        Frm_Medicos_Acreditados.DGV.Columns(16).Width = 70
        Frm_Medicos_Acreditados.DGV.Columns(16).Visible = True
        Frm_Medicos_Acreditados.DGV.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Medicos_Acreditados.DGV.Columns(17).HeaderText = "Fecha Desde"
        Frm_Medicos_Acreditados.DGV.Columns(17).Width = 70
        Frm_Medicos_Acreditados.DGV.Columns(17).Visible = True
        Frm_Medicos_Acreditados.DGV.Columns(17).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Medicos_Acreditados.DGV.Columns(18).HeaderText = "Fecha Hasta"
        Frm_Medicos_Acreditados.DGV.Columns(18).Width = 70
        Frm_Medicos_Acreditados.DGV.Columns(18).Visible = True
        Frm_Medicos_Acreditados.DGV.Columns(18).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Medicos_Acreditados.DGV.Columns(19).HeaderText = "FACEBOOK"
        Frm_Medicos_Acreditados.DGV.Columns(19).Width = 0
        Frm_Medicos_Acreditados.DGV.Columns(19).Visible = False

        Frm_Medicos_Acreditados.DGV.Columns(20).HeaderText = "TWITTER"
        Frm_Medicos_Acreditados.DGV.Columns(20).Width = 0
        Frm_Medicos_Acreditados.DGV.Columns(20).Visible = False

        Frm_Medicos_Acreditados.DGV.Columns(21).HeaderText = "Whatsapp"
        Frm_Medicos_Acreditados.DGV.Columns(21).Width = 80
        Frm_Medicos_Acreditados.DGV.Columns(21).Visible = True
        Frm_Medicos_Acreditados.DGV.Columns(21).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Medicos_Acreditados.DGV.Columns(22).HeaderText = "Observaciones"
        Frm_Medicos_Acreditados.DGV.Columns(22).Width = 150
        Frm_Medicos_Acreditados.DGV.Columns(22).Visible = True
        Frm_Medicos_Acreditados.DGV.Columns(22).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Medicos_Acreditados.DGV.Columns(23).HeaderText = "AN"
        Frm_Medicos_Acreditados.DGV.Columns(23).Width = 0
        Frm_Medicos_Acreditados.DGV.Columns(23).Visible = False

        Frm_Medicos_Acreditados.DGV.Columns(24).HeaderText = "FOTO"
        Frm_Medicos_Acreditados.DGV.Columns(24).Width = 0
        Frm_Medicos_Acreditados.DGV.Columns(24).Visible = False

    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        Dim I As Integer
        For I = 0 To Frm_Medicos_Acreditados.DGV.RowCount - 1
            If Frm_Medicos_Acreditados.DGV.Item(0, I).Value = vMAc_ID_MA Then
                Frm_Medicos_Acreditados.DGV.Rows(I).Selected = True
                Frm_Medicos_Acreditados.DGV.CurrentCell = Frm_Medicos_Acreditados.DGV.Item(0, I)
                Exit Sub
            End If
        Next
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
End Class