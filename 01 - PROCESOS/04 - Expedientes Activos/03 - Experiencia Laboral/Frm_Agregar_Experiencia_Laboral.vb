Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Public Class Frm_Agregar_Experiencia_Laboral
    Private Sub Frm_Agregar_Experiencia_Laboral_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = ""
            Me.TextBox2.Text = ""
            Me.TextBox3.Text = ""
            Me.TextBox4.Text = "0"
            Me.TextBox5.Text = ""
            Me.TextBox6.Text = ""
            Me.TextBox7.Text = ""
            Me.DateTimePicker1.Value = Now
            Me.DateTimePicker2.Value = Now
            Me.TextBox8.Text = "0"
            Me.TextBox9.Text = ""
            Me.TextBox10.Text = ""
            Me.TextBox2.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Agregar_Experiencia_Laboral_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = "0"
        Me.TextBox5.Text = ""
        Me.TextBox6.Text = ""
        Me.TextBox7.Text = ""
        Me.DateTimePicker1.Value = Now
        Me.DateTimePicker2.Value = Now
        Me.TextBox8.Text = "0"
        Me.TextBox9.Text = ""
        Me.TextBox10.Text = ""
        Me.TextBox2.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = "0"
        Me.TextBox5.Text = ""
        Me.TextBox6.Text = ""
        Me.TextBox7.Text = ""
        Me.DateTimePicker1.Value = Now
        Me.DateTimePicker2.Value = Now
        Me.TextBox8.Text = "0"
        Me.TextBox9.Text = ""
        Me.TextBox10.Text = ""
        Me.TextBox2.Focus()
        CERRAR = True
        Me.Close()
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
    Private Sub GENERAR_CODIGO()
        CODIGO = 0
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            MiDataAdapter = New SqlDataAdapter("SELECT ID_EL FROM [MAESTRO DE EXPERIENCIA LABORAL] ORDER BY ID_EL", Conexion.CxRRHH)
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
    Dim vcMIXTA, vcMILITAR, vcPAME As Boolean
    Private Sub BUSCAR_ESTRUCTURA_CARGO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [MAESTRO DE CARGOS] WHERE ID_M_C = '" & Frm_Expedientes_Activos.TextBox36.Text & "'", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                vcMIXTA = MiDataTable.Rows(0).Item(15).ToString
                vcMILITAR = MiDataTable.Rows(0).Item(16).ToString
                vcPAME = MiDataTable.Rows(0).Item(17).ToString
            Else
                vcMIXTA = False
                vcMILITAR = False
                vcPAME = False
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
        Me.TextBox1.Text = "0"
        Call GENERAR_CODIGO()
        If TextBox1.Text = "0" Or TextBox1.Text = "" Then
            MsgBox("No se ha generado el Codigo del registro", vbInformation, "Mensaje del Sistema")
            Me.TextBox1.Text = "0"
            Me.Button2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox2.Text) = 0 Then
            MsgBox("Debe digitar el Nombre de la Empresa correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox3.Text) = 0 Then
            MsgBox("Debe digitar la Dirección de la Empresa correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox3.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox4.Text) = 0 Then
            Me.TextBox4.Text = ""
        End If
        If Len(Me.TextBox5.Text) = 0 Then
            MsgBox("Debe digitar la Actividad de la Empresa correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox5.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox6.Text) = 0 Then
            MsgBox("Debe digitar el Cargo Desempeñado en la Empresa correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox6.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox7.Text) = 0 Then
            MsgBox("Debe digitar el Nombre del Jefe Inmediato correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox7.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox8.Text) = 0 Or Me.TextBox8.Text = "0" Then
            MsgBox("Debe digitar el Salario devengado correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox8.Focus()
        End If
        If Len(Me.TextBox9.Text) = 0 Then
            MsgBox("Debe digitar los Motivos de retiro correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox9.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox10.Text) = 0 Then
            MsgBox("Debe digitar las Obligaciones y Responsabilidades correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox10.Focus()
            Exit Sub
        End If
        Call GRABAR_REGISTRO()
        vmelID_EL = Me.TextBox1.Text
        CADENAsql = "SELECT * FROM [MAESTRO DE EXPERIENCIA LABORAL] WHERE ID_M_P = " & Frm_Expedientes_Activos.TextBox35.Text & " ORDER BY ID_M_P, ID_EL"
        Call CARGAR_DATAGRIDVIEW_2()
        Call ARMAR_DATAGRIDVIEW_2()
        Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_2()
        CERRAR = False
    End Sub
    Private Sub GRABAR_REGISTRO()
        Dim fFEC_INI As String = Me.DateTimePicker1.Text
        Dim fFEC_FIN As String = Me.DateTimePicker2.Text
        Dim fechaInicial = "Convert(VARCHAR(10), Convert(Date, '" + fFEC_INI + "', 105), 23)"
        Dim fechaFinal = "Convert(VARCHAR(10), Convert(Date, '" + fFEC_FIN + "', 105), 23)"

        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE EXPERIENCIA LABORAL] (ID_EL, ID_M_P, EMPRESA_NOMBRE, EMPRESA_DIRECCION, EMPRESA_TELEFONO, EMPRESA_ACTIVIDAD, CARGO_DESEM, JEFE_INMEDIATO, SALARIO, FEC_INI, FEC_FIN, CAUSA_RETIRO, OBLI_RESP)" &
                                  "values (" & CInt(Me.TextBox1.Text) & ", " & Frm_Expedientes_Activos.TextBox35.Text & ", '" & Me.TextBox2.Text & "', '" & Me.TextBox3.Text & "', '" & Me.TextBox4.Text & "', '" & Me.TextBox5.Text & "', '" & Me.TextBox6.Text & "', '" & Me.TextBox7.Text & "', '" & Me.TextBox8.Text & "', " & fechaInicial & ", " & fechaFinal & ", '" & Me.TextBox9.Text & "', '" & Me.TextBox10.Text & "')"
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
        MiAdaptador.Fill(MiDataSet, "[MAESTRO DE EXPERIENCIA LABORAL]")
        Frm_Expedientes_Activos.DGV2.DataSource = MiDataSet.Tables("[MAESTRO DE EXPERIENCIA LABORAL]")
        Frm_Expedientes_Activos.Label80.Text = Frm_Expedientes_Activos.DGV2.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_2()
        Frm_Expedientes_Activos.DGV2.Columns(0).HeaderText = "Id"
        Frm_Expedientes_Activos.DGV2.Columns(0).Width = 30
        Frm_Expedientes_Activos.DGV2.Columns(0).Visible = True
        Frm_Expedientes_Activos.DGV2.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV2.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Frm_Expedientes_Activos.DGV2.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Frm_Expedientes_Activos.DGV2.Columns(1).HeaderText = "ID_M_P"
        Frm_Expedientes_Activos.DGV2.Columns(1).Width = 0
        Frm_Expedientes_Activos.DGV2.Columns(1).Visible = False

        Frm_Expedientes_Activos.DGV2.Columns(2).HeaderText = "Nombre de Empresa"
        Frm_Expedientes_Activos.DGV2.Columns(2).Width = 200
        Frm_Expedientes_Activos.DGV2.Columns(2).Visible = True
        Frm_Expedientes_Activos.DGV2.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV2.Columns(2).Frozen = True

        Frm_Expedientes_Activos.DGV2.Columns(3).HeaderText = "Direccion de Empresa"
        Frm_Expedientes_Activos.DGV2.Columns(3).Width = 200
        Frm_Expedientes_Activos.DGV2.Columns(3).Visible = True
        Frm_Expedientes_Activos.DGV2.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV2.Columns(4).HeaderText = "Telefono de Empresa"
        Frm_Expedientes_Activos.DGV2.Columns(4).Width = 110
        Frm_Expedientes_Activos.DGV2.Columns(4).Visible = True
        Frm_Expedientes_Activos.DGV2.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV2.Columns(5).HeaderText = "Actividad de Empresa"
        Frm_Expedientes_Activos.DGV2.Columns(5).Width = 200
        Frm_Expedientes_Activos.DGV2.Columns(5).Visible = True
        Frm_Expedientes_Activos.DGV2.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV2.Columns(6).HeaderText = "Cargo que Desempeñado"
        Frm_Expedientes_Activos.DGV2.Columns(6).Width = 200
        Frm_Expedientes_Activos.DGV2.Columns(6).Visible = True
        Frm_Expedientes_Activos.DGV2.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV2.Columns(7).HeaderText = "Jefe Inmediato"
        Frm_Expedientes_Activos.DGV2.Columns(7).Width = 200
        Frm_Expedientes_Activos.DGV2.Columns(7).Visible = True
        Frm_Expedientes_Activos.DGV2.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV2.Columns(8).HeaderText = "Salario"
        Frm_Expedientes_Activos.DGV2.Columns(8).Width = 70
        Frm_Expedientes_Activos.DGV2.Columns(8).Visible = True
        Frm_Expedientes_Activos.DGV2.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Frm_Expedientes_Activos.DGV2.Columns(8).DefaultCellStyle.Format = "##,##0.00"

        Frm_Expedientes_Activos.DGV2.Columns(9).HeaderText = "Desde"
        Frm_Expedientes_Activos.DGV2.Columns(9).Width = 70
        Frm_Expedientes_Activos.DGV2.Columns(9).Visible = True
        Frm_Expedientes_Activos.DGV2.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV2.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Frm_Expedientes_Activos.DGV2.Columns(10).HeaderText = "Hasta"
        Frm_Expedientes_Activos.DGV2.Columns(10).Width = 70
        Frm_Expedientes_Activos.DGV2.Columns(10).Visible = True
        Frm_Expedientes_Activos.DGV2.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV2.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Frm_Expedientes_Activos.DGV2.Columns(11).HeaderText = "Causa de Retiro"
        Frm_Expedientes_Activos.DGV2.Columns(11).Width = 200
        Frm_Expedientes_Activos.DGV2.Columns(11).Visible = True
        Frm_Expedientes_Activos.DGV2.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV2.Columns(12).HeaderText = "Obligaciones y Responsabildiades"
        Frm_Expedientes_Activos.DGV2.Columns(12).Width = 200
        Frm_Expedientes_Activos.DGV2.Columns(12).Visible = True
        Frm_Expedientes_Activos.DGV2.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_2()
        Dim I As Integer
        For I = 0 To Frm_Expedientes_Activos.DGV2.RowCount - 1
            If Frm_Expedientes_Activos.DGV2.Item(0, I).Value = vmelID_EL Then
                Frm_Expedientes_Activos.DGV2.Rows(I).Selected = True
                Frm_Expedientes_Activos.DGV2.CurrentCell = Frm_Expedientes_Activos.DGV2.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
End Class