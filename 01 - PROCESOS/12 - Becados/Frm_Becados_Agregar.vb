Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Becados_Agregar
    Private Sub Frm_Becados_Agregar_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox4.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Becados_Agregar_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Call INICIAR_OBJETOS()
        Me.TextBox4.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox4.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub INICIAR_OBJETOS()
        Me.TextBox1.Text = "0"
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Me.TextBox6.Text = ""
        Me.TextBox7.Text = ""
        Call CARGAR_COMBO_CAT_DE_PAIS()
        Me.TextBox18.Text = ""
        Me.TextBox8.Text = ""
        Me.DateTimePicker2.Value = Now
        Me.DateTimePicker3.Value = Now
        Me.TextBox9.Text = ""
        Me.TextBox10.Text = ""
        Me.TextBox11.Text = ""
        Me.TextBox12.Text = "0"
        Me.TextBox13.Text = "0"
        Me.TextBox14.Text = "0"
        Me.TextBox15.Text = "0"
        Me.TextBox16.Text = "0"
        Me.TextBox20.Text = "0"
        Me.TextBox17.Text = ""
        Me.TextBox19.Text = "0"
        Me.TextBox4.Focus()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_PAIS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE PAIS] ORDER BY ID_PAIS", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
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
    Dim cEMPLEADO As String
    Dim DATO As Boolean
    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            cEMPLEADO = CInt(Val(TextBox4.Text))
            TextBox4.Text = Format(CInt(cEMPLEADO), "0000000000")
            If Len(Me.TextBox4.Text) = 0 Or (Me.TextBox4.Text = "0000000000") Then
                MsgBox("El código que se digitó no es válido", vbInformation, "Mensaje del Sistema")
                Me.TextBox1.Focus()
                Exit Sub
            Else
                Call BUSCAR_DATO()
                If DATO = True Then
                    e.Handled = True
                    SendKeys.Send("{TAB}")
                Else
                    MsgBox("El código que se digitó no es válido", vbInformation, "Mensaje del Sistema")
                    Me.TextBox4.SelectAll()
                    Me.TextBox4.Focus()
                    Exit Sub
                End If
            End If
        End If
    End Sub
    Private Sub BUSCAR_DATO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT ID_M_P, CODIGO, APELLIDOS, NOMBRES, N_CEDULA, N_MINSA, D_T_SEXO, ID_ESTADO_P FROM [dbo].[VISTA MAESTRO DE PERSONAS] WHERE CODIGO = '" & Me.TextBox4.Text & "' AND ID_ESTADO_P = 1", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                Me.TextBox19.Text = MiDataTable.Rows(0).Item(0).ToString
                Me.TextBox2.Text = MiDataTable.Rows(0).Item(2).ToString
                Me.TextBox3.Text = MiDataTable.Rows(0).Item(3).ToString
                Me.TextBox5.Text = MiDataTable.Rows(0).Item(4).ToString
                Me.TextBox6.Text = MiDataTable.Rows(0).Item(5).ToString
                Me.TextBox7.Text = MiDataTable.Rows(0).Item(6).ToString
                DATO = True
            Else
                Me.TextBox1.Text = ""
                Me.TextBox2.Text = ""
                Me.TextBox3.Text = ""
                Me.TextBox5.Text = ""
                Me.TextBox6.Text = ""
                Me.TextBox7.Text = ""
                DATO = False
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
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
    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox18_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox18.KeyDown
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
    Private Sub TextBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox9.KeyDown
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
    Private Sub TextBox15_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox15.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox16_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox16.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox17_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox17.KeyDown
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
            MiDataAdapter = New SqlDataAdapter("SELECT ID_B FROM [MAESTRO DE BECADOS] ORDER BY ID_B", Conexion.CxRRHH)
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
        If TextBox19.Text = "0" Or TextBox19.Text = "" Then
            MsgBox("No se encuentra el Codigo del Empleado", vbInformation, "Mensaje del Sistema")
            Me.TextBox19.Text = "0"
            Me.Button2.Focus()
            Exit Sub
        End If
        Me.TextBox1.Text = "0"
        Call GENERAR_CODIGO()
        If TextBox1.Text = "0" Or TextBox1.Text = "" Then
            MsgBox("No se ha generado el Codigo del registro", vbInformation, "Mensaje del Sistema")
            Me.TextBox1.Text = "0"
            Me.Button2.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Pais de Estudio Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If TextBox18.Text = "" Then
            MsgBox("Debe Digitar el Nombre del Estudio Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox18.Focus()
            Exit Sub
        End If
        If TextBox8.Text = "" Then
            MsgBox("Debe Digitar el Centro de Estudio Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox8.Focus()
            Exit Sub
        End If
        If TextBox9.Text = "" Then
            TextBox9.Text = "0"
        End If
        If TextBox10.Text = "" Then
            TextBox10.Text = "."
        End If
        If TextBox11.Text = "" Then
            TextBox11.Text = "."
        End If
        If TextBox12.Text = "" Then
            TextBox12.Text = "0"
        End If
        If TextBox13.Text = "" Then
            TextBox13.Text = "0"
        End If
        If TextBox14.Text = "" Then
            TextBox14.Text = "0"
        End If
        If TextBox15.Text = "" Then
            TextBox15.Text = "0"
        End If
        If TextBox16.Text = "" Then
            TextBox16.Text = "0"
        End If
        If TextBox20.Text = "" Then
            TextBox20.Text = "0"
        End If
        If TextBox17.Text = "" Then
            TextBox17.Text = "."
        End If
        TOT = CDbl(Me.TextBox12.Text) + CDbl(Me.TextBox13.Text) + CDbl(Me.TextBox14.Text) + CDbl(Me.TextBox15.Text) + CDbl(Me.TextBox16.Text) + CDbl(Me.TextBox20.Text)
        Call Me.GRABAR_REGISTRO()
        CERRAR = False
        mbID_B = Me.TextBox1.Text
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
        Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        Call INICIAR_OBJETOS()
        Me.TextBox4.Focus()
    End Sub
    Dim TOT As String
    Private Sub GRABAR_REGISTRO()
        Dim fFEC_1 As String = Me.DateTimePicker2.Text
        Dim fFEC_2 As String = Me.DateTimePicker3.Text
        Dim F1 = "Convert(VARCHAR(10), Convert(Date, '" + fFEC_1 + "', 105), 23)"
        Dim F2 = "Convert(VARCHAR(10), Convert(Date, '" + fFEC_2 + "', 105), 23)"

        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim SQL As String
        Try
            SQL = "INSERT INTO [MAESTRO DE BECADOS] (ID_B, ID_M_P, NOMBRE_ESTUDIO, ID_PAIS, CENTRO_ESTUDIO, FEC_INI, FEC_FIN, TELEFONO, NO_CUENTA, CORREO, MATRICULA, INSCRIPCION, ALIMENTACION, HOSPEDAJE, PAGO_TITULO, OTROS, TOTAL, OBSERVACIONES)" &
                  "values (" & CInt(Me.TextBox1.Text) & ",  " & CInt(Me.TextBox19.Text) & ", '" & Me.TextBox18.Text & "', " & Me.ComboBox1.SelectedValue & ",  '" & Me.TextBox8.Text & "', " & F1 & ", " & F2 & ", '" & Me.TextBox9.Text & "', '" & Me.TextBox10.Text & "', '" & Me.TextBox11.Text & "', '" & Me.TextBox12.Text & "', '" & Me.TextBox13.Text & "',  '" & Me.TextBox14.Text & "', '" & Me.TextBox15.Text & "', '" & Me.TextBox16.Text & "', '" & Me.TextBox20.Text & "', '" & TOT & "', '" & Me.TextBox17.Text & "')"
            Dim cmd As New SqlCommand(SQL, CxRRHH)
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
        Dim MiAdaptador As New SqlClient.SqlDataAdapter("Select * from [VISTA MAESTRO DE BECADOS] ORDER BY NOMBRES, APELLIDOS", Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE BECADOS]")
        Frm_Becados.DGV.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE BECADOS]")
        Frm_Becados.Label2.Text = "Total de Registros: " & Frm_Becados.DGV.RowCount
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW()
        Frm_Becados.DGV.Columns(0).HeaderText = "Id"
        Frm_Becados.DGV.Columns(0).Width = 30
        Frm_Becados.DGV.Columns(0).Visible = True
        Frm_Becados.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Becados.DGV.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Frm_Becados.DGV.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Frm_Becados.DGV.Columns(1).HeaderText = "ID_M_P"
        Frm_Becados.DGV.Columns(1).Width = 0
        Frm_Becados.DGV.Columns(1).Visible = False
        Frm_Becados.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Becados.DGV.Columns(2).HeaderText = "Nombres"
        Frm_Becados.DGV.Columns(2).Width = 140
        Frm_Becados.DGV.Columns(2).Visible = True
        Frm_Becados.DGV.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Becados.DGV.Columns(2).DefaultCellStyle.BackColor = Color.Khaki

        Frm_Becados.DGV.Columns(3).HeaderText = "Apellidos"
        Frm_Becados.DGV.Columns(3).Width = 140
        Frm_Becados.DGV.Columns(3).Visible = True
        Frm_Becados.DGV.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Becados.DGV.Columns(3).DefaultCellStyle.BackColor = Color.Khaki

        Frm_Becados.DGV.Columns(4).HeaderText = "Codigo"
        Frm_Becados.DGV.Columns(4).Width = 70
        Frm_Becados.DGV.Columns(4).Visible = True
        Frm_Becados.DGV.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Becados.DGV.Columns(5).HeaderText = "Cedula"
        Frm_Becados.DGV.Columns(5).Width = 0
        Frm_Becados.DGV.Columns(5).Visible = False
        Frm_Becados.DGV.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Becados.DGV.Columns(6).HeaderText = "No. MINSA"
        Frm_Becados.DGV.Columns(6).Width = 0
        Frm_Becados.DGV.Columns(6).Visible = False
        Frm_Becados.DGV.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        '
        Frm_Becados.DGV.Columns(7).HeaderText = "D_T_SEXO"
        Frm_Becados.DGV.Columns(7).Width = 0
        Frm_Becados.DGV.Columns(7).Visible = False

        Frm_Becados.DGV.Columns(8).HeaderText = "NOMBRE_ESTUDIO"
        Frm_Becados.DGV.Columns(8).Width = 0
        Frm_Becados.DGV.Columns(8).Visible = False

        Frm_Becados.DGV.Columns(9).HeaderText = "ID_PAIS"
        Frm_Becados.DGV.Columns(9).Width = 0
        Frm_Becados.DGV.Columns(9).Visible = False

        Frm_Becados.DGV.Columns(10).HeaderText = "D_PAIS"
        Frm_Becados.DGV.Columns(10).Width = 0
        Frm_Becados.DGV.Columns(10).Visible = False

        Frm_Becados.DGV.Columns(11).HeaderText = "CENTRO_ESTUDIO"
        Frm_Becados.DGV.Columns(11).Width = 0
        Frm_Becados.DGV.Columns(11).Visible = False

        Frm_Becados.DGV.Columns(12).HeaderText = "Fecha Inicio"
        Frm_Becados.DGV.Columns(12).Width = 70
        Frm_Becados.DGV.Columns(12).Visible = True
        Frm_Becados.DGV.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Becados.DGV.Columns(13).HeaderText = "Fecha Fin"
        Frm_Becados.DGV.Columns(13).Width = 70
        Frm_Becados.DGV.Columns(13).Visible = True
        Frm_Becados.DGV.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Becados.DGV.Columns(14).HeaderText = "TELEFONO"
        Frm_Becados.DGV.Columns(14).Width = 0
        Frm_Becados.DGV.Columns(14).Visible = False

        Frm_Becados.DGV.Columns(15).HeaderText = "No. Cuenta"
        Frm_Becados.DGV.Columns(15).Width = 0
        Frm_Becados.DGV.Columns(15).Visible = False
        Frm_Becados.DGV.Columns(15).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Becados.DGV.Columns(16).HeaderText = "Correo"
        Frm_Becados.DGV.Columns(16).Width = 0
        Frm_Becados.DGV.Columns(16).Visible = False
        Frm_Becados.DGV.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Becados.DGV.Columns(17).HeaderText = "Matricula"
        Frm_Becados.DGV.Columns(17).Width = 73
        Frm_Becados.DGV.Columns(17).Visible = True
        Frm_Becados.DGV.Columns(17).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Becados.DGV.Columns(17).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Frm_Becados.DGV.Columns(18).HeaderText = "Inscripcion"
        Frm_Becados.DGV.Columns(18).Width = 73
        Frm_Becados.DGV.Columns(18).Visible = True
        Frm_Becados.DGV.Columns(18).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Becados.DGV.Columns(18).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Frm_Becados.DGV.Columns(19).HeaderText = "Alimentacion"
        Frm_Becados.DGV.Columns(19).Width = 73
        Frm_Becados.DGV.Columns(19).Visible = True
        Frm_Becados.DGV.Columns(19).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Becados.DGV.Columns(19).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Frm_Becados.DGV.Columns(20).HeaderText = "Hospedaje"
        Frm_Becados.DGV.Columns(20).Width = 73
        Frm_Becados.DGV.Columns(20).Visible = True
        Frm_Becados.DGV.Columns(20).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Becados.DGV.Columns(20).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Frm_Becados.DGV.Columns(21).HeaderText = "Pago Titulo"
        Frm_Becados.DGV.Columns(21).Width = 73
        Frm_Becados.DGV.Columns(21).Visible = True
        Frm_Becados.DGV.Columns(21).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Becados.DGV.Columns(21).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Frm_Becados.DGV.Columns(22).HeaderText = "Otros"
        Frm_Becados.DGV.Columns(22).Width = 73
        Frm_Becados.DGV.Columns(22).Visible = True
        Frm_Becados.DGV.Columns(22).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Becados.DGV.Columns(22).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Frm_Becados.DGV.Columns(23).HeaderText = "Total"
        Frm_Becados.DGV.Columns(23).Width = 73
        Frm_Becados.DGV.Columns(23).Visible = True
        Frm_Becados.DGV.Columns(23).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Becados.DGV.Columns(23).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Frm_Becados.DGV.Columns(24).HeaderText = "OBSERVACIONES"
        Frm_Becados.DGV.Columns(24).Width = 0
        Frm_Becados.DGV.Columns(24).Visible = False

    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        Dim I As Integer
        For I = 0 To Frm_Becados.DGV.RowCount - 1
            If Frm_Becados.DGV.Item(0, I).Value = mbID_B Then
                Frm_Becados.DGV.Rows(I).Selected = True
                Frm_Becados.DGV.CurrentCell = Frm_Becados.DGV.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub TextBox20_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox20.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
End Class