Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Dosimetria_Add
    Private Sub Frm_Dosimetria_Add_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox2.Focus()
            CERRAR = True
            Me.TextBox2.Focus()
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Dosimetria_Add_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = "0"
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Me.TextBox7.Text = ""
        Me.TextBox6.Text = ""
        Me.DateTimePicker1.Value = Now
        Me.DateTimePicker2.Value = Now
        Me.DateTimePicker10.Value = Now
        Me.DateTimePicker20.Value = Now
        Me.DateTimePicker3.Value = Now
        Call CARGAR_COMBO_CAT_DE_PERIODO_DOSIMETRO()
        Me.TextBox4.Text = ""
        Me.TextBox2.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox2.Focus()
        CERRAR = True
        Me.TextBox2.Focus()
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_PERIODO_DOSIMETRO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE PERIODO DOSIMETRO] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_PER_DOSIME"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
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
            MiDataAdapter = New SqlDataAdapter("SELECT ID_DOS FROM [MAESTRO DE HIGIENE DOSIMETRIA] ORDER BY ID_DOS", Conexion.CxRRHH)
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
        If Len(Me.TextBox2.Text) = 0 Then
            Me.TextBox2.Text = "0"
        End If
        If Len(Me.TextBox3.Text) = 0 Then
            Me.TextBox3.Text = "0"
        End If
        If Len(Me.TextBox5.Text) = 0 Then
            Me.TextBox5.Text = "0"
        End If
        If Len(Me.TextBox7.Text) = 0 Then
            Me.TextBox7.Text = "0"
        End If
        If Len(Me.TextBox6.Text) = 0 Then
            Me.TextBox6.Text = "0"
        End If
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Periodo Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox4.Text) = 0 Then
            Me.TextBox4.Text = "."
        End If
        Call GRABAR_REGISTRO()
        h5ID_DOS = Me.TextBox1.Text
        CADENAsql = "Select * from [dbo].[VISTA MAESTRO DE HIGIENE DOSIMETRIA] WHERE ID_M_P = " & Frm_Dosimetria.TextBox36.Text & " ORDER BY FECHA_LECTURA"
        Call Me.CARGAR_DATAGRIDVIEW()
        Call Me.ARMAR_DATAGRIDVIEW()
        Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        Me.TextBox1.Text = "0"
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Me.TextBox7.Text = ""
        Me.TextBox6.Text = ""
        Me.DateTimePicker1.Value = Now
        Me.DateTimePicker2.Value = Now
        Me.DateTimePicker10.Value = Now
        Me.DateTimePicker20.Value = Now
        Me.TextBox4.Text = ""
        Me.TextBox2.Focus()
        CERRAR = False
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        Dim I As Integer
        For I = 0 To Frm_Dosimetria.DGV.RowCount - 1
            If Frm_Dosimetria.DGV.Item(0, I).Value = h5ID_DOS Then
                Frm_Dosimetria.DGV.Rows(I).Selected = True
                Frm_Dosimetria.DGV.CurrentCell = Frm_Dosimetria.DGV.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE HIGIENE DOSIMETRIA]")
        Frm_Dosimetria.DGV.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE HIGIENE DOSIMETRIA]")
        Frm_Dosimetria.Label2.Text = "Total de Registros: " & Frm_Dosimetria.DGV.RowCount
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW()
        Frm_Dosimetria.DGV.Columns(0).HeaderText = "Id"
        Frm_Dosimetria.DGV.Columns(0).Width = 30
        Frm_Dosimetria.DGV.Columns(0).Visible = True
        Frm_Dosimetria.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Dosimetria.DGV.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Frm_Dosimetria.DGV.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Frm_Dosimetria.DGV.Columns(1).HeaderText = "Departamento"
        Frm_Dosimetria.DGV.Columns(1).Width = 190
        Frm_Dosimetria.DGV.Columns(1).Visible = True
        Frm_Dosimetria.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Dosimetria.DGV.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Frm_Dosimetria.DGV.Columns(2).HeaderText = "Servicio/ Area"
        Frm_Dosimetria.DGV.Columns(2).Width = 180
        Frm_Dosimetria.DGV.Columns(2).Visible = True
        Frm_Dosimetria.DGV.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Dosimetria.DGV.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Frm_Dosimetria.DGV.Columns(3).HeaderText = "Cargo"
        Frm_Dosimetria.DGV.Columns(3).Width = 100
        Frm_Dosimetria.DGV.Columns(3).Visible = True
        Frm_Dosimetria.DGV.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Dosimetria.DGV.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Frm_Dosimetria.DGV.Columns(4).HeaderText = "ID_M_P"
        Frm_Dosimetria.DGV.Columns(4).Width = 0
        Frm_Dosimetria.DGV.Columns(4).Visible = False

        Frm_Dosimetria.DGV.Columns(5).HeaderText = "Codigo Dosimetro"
        Frm_Dosimetria.DGV.Columns(5).Width = 70
        Frm_Dosimetria.DGV.Columns(5).Visible = True
        Frm_Dosimetria.DGV.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Dosimetria.DGV.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Dosimetria.DGV.Columns(6).HeaderText = "Fecha Lectura"
        Frm_Dosimetria.DGV.Columns(6).Width = 70
        Frm_Dosimetria.DGV.Columns(6).Visible = True
        Frm_Dosimetria.DGV.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Dosimetria.DGV.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Frm_Dosimetria.DGV.Columns(7).HeaderText = "FECHA_PE_I"
        Frm_Dosimetria.DGV.Columns(7).Width = 0
        Frm_Dosimetria.DGV.Columns(7).Visible = False
        Frm_Dosimetria.DGV.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Dosimetria.DGV.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Frm_Dosimetria.DGV.Columns(7).DefaultCellStyle.BackColor = Color.Khaki

        Frm_Dosimetria.DGV.Columns(8).HeaderText = "FECHA_PE_F"
        Frm_Dosimetria.DGV.Columns(8).Width = 0
        Frm_Dosimetria.DGV.Columns(8).Visible = False
        Frm_Dosimetria.DGV.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Dosimetria.DGV.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Frm_Dosimetria.DGV.Columns(8).DefaultCellStyle.BackColor = Color.Khaki

        Frm_Dosimetria.DGV.Columns(9).HeaderText = "Dosis Eq. Per."
        Frm_Dosimetria.DGV.Columns(9).Width = 90
        Frm_Dosimetria.DGV.Columns(9).Visible = True
        Frm_Dosimetria.DGV.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Dosimetria.DGV.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Frm_Dosimetria.DGV.Columns(9).DefaultCellStyle.BackColor = Color.Khaki

        Frm_Dosimetria.DGV.Columns(10).HeaderText = "Dosis Acum. I S"
        Frm_Dosimetria.DGV.Columns(10).Width = 90
        Frm_Dosimetria.DGV.Columns(10).Visible = True
        Frm_Dosimetria.DGV.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Dosimetria.DGV.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Frm_Dosimetria.DGV.Columns(11).HeaderText = "Dosis Acum. II S"
        Frm_Dosimetria.DGV.Columns(11).Width = 90
        Frm_Dosimetria.DGV.Columns(11).Visible = True
        Frm_Dosimetria.DGV.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Dosimetria.DGV.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Frm_Dosimetria.DGV.Columns(12).HeaderText = "Dosis Acum. Anual"
        Frm_Dosimetria.DGV.Columns(12).Width = 90
        Frm_Dosimetria.DGV.Columns(12).Visible = True
        Frm_Dosimetria.DGV.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Dosimetria.DGV.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Frm_Dosimetria.DGV.Columns(12).DefaultCellStyle.BackColor = Color.Khaki

        Frm_Dosimetria.DGV.Columns(13).HeaderText = ""
        Frm_Dosimetria.DGV.Columns(13).Width = 0
        Frm_Dosimetria.DGV.Columns(13).Visible = False

        Frm_Dosimetria.DGV.Columns(14).HeaderText = ""
        Frm_Dosimetria.DGV.Columns(14).Width = 0
        Frm_Dosimetria.DGV.Columns(14).Visible = False

        Frm_Dosimetria.DGV.Columns(15).HeaderText = ""
        Frm_Dosimetria.DGV.Columns(15).Width = 0
        Frm_Dosimetria.DGV.Columns(15).Visible = False

        Frm_Dosimetria.DGV.Columns(16).HeaderText = "Periodo"
        Frm_Dosimetria.DGV.Columns(16).Width = 70
        Frm_Dosimetria.DGV.Columns(16).Visible = True
        Frm_Dosimetria.DGV.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Dosimetria.DGV.Columns(16).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Frm_Dosimetria.DGV.Columns(16).DefaultCellStyle.BackColor = Color.Khaki

        Frm_Dosimetria.DGV.Columns(17).HeaderText = "Observaciones"
        Frm_Dosimetria.DGV.Columns(17).Width = 0
        Frm_Dosimetria.DGV.Columns(17).Visible = False
        Frm_Dosimetria.DGV.Columns(17).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Dosimetria.DGV.Columns(17).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Frm_Dosimetria.DGV.Columns(17).DefaultCellStyle.BackColor = Color.Khaki
    End Sub
    Private Sub GRABAR_REGISTRO()
        Dim fFECHA_1 As String
        fFECHA_1 = Mid(Me.DateTimePicker1.Value, 1, 10)
        Dim f1 As Object = If(fFECHA_1 <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_1 + "', 105), 23)", "NULL")
        Dim fFECHA_2 As String
        fFECHA_2 = Mid(Me.DateTimePicker10.Value, 1, 10)
        Dim f2 As Object = If(fFECHA_2 <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_2 + "', 105), 23)", "NULL")
        Dim fFECHA_3 As String
        fFECHA_3 = Mid(Me.DateTimePicker20.Value, 1, 10)
        Dim f3 As Object = If(fFECHA_3 <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_3 + "', 105), 23)", "NULL")
        Dim fFECHA_4 As String
        fFECHA_4 = Mid(Me.DateTimePicker2.Value, 1, 10)
        Dim f4 As Object = If(fFECHA_4 <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_4 + "', 105), 23)", "NULL")
        Dim fFECHA_5 As String
        fFECHA_5 = Mid(Me.DateTimePicker3.Value, 1, 10)
        Dim f5 As Object = If(fFECHA_4 <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_5 + "', 105), 23)", "NULL")

        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE HIGIENE DOSIMETRIA] (ID_DOS, DEPARTAMENTO, SERVICIO, CARGO, ID_M_P, CODIGO_DOSIMETRO, FECHA_LECTURA, FECHA_PE_I, FECHA_PE_F, DOSIS_EP, DOSIS_APS, DOSIS_ASS, DOSIS_AA, FECHA_I_PORT, FECHA_C_PROGRAMADA, ID_PER_DOSIME, OBSERVACION)" &
            "values (" & CInt(Me.TextBox1.Text) & ", '" & Frm_Dosimetria.TextBox6.Text & "', '" & Frm_Dosimetria.TextBox7.Text & "', '" & Frm_Dosimetria.TextBox4.Text & "', " & Frm_Dosimetria.TextBox36.Text & ", '" & Me.TextBox2.Text & "', " & f1 & ", " & f2 & ", " & f3 & ", '" & Me.TextBox3.Text & "', '" & Me.TextBox5.Text & "', '" & Me.TextBox7.Text & "', '" & Me.TextBox6.Text & "', " & f5 & ", " & f4 & ", " & Me.ComboBox1.SelectedValue & ", '" & Me.TextBox4.Text & "')"
            MiSqlCommand.ExecuteNonQuery()
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
    Private Sub DateTimePicker10_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker10.KeyDown
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
    Private Sub DateTimePicker2_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker2.KeyDown
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
    Private Sub DateTimePicker20_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker20.KeyDown
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
            Me.TextBox6.Text = Val(Me.TextBox5.Text) + Val(Me.TextBox7.Text)
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