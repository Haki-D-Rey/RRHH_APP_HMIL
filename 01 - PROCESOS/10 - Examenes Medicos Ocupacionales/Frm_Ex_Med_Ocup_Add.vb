Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Ex_Med_Ocup_Add
    Private Sub Frm_Ex_Med_Ocup_Add_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            CERRAR = True
            Me.DateTimePicker1.Focus()
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Ex_Med_Ocup_Add_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = "0"
        Me.DateTimePicker1.Value = Now
        Me.TextBox2.Text = ""
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_EVALUACION()
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = Frm_Ex_Med_Ocup.TextBox4.Text
        Me.TextBox6.Text = ""
        Call BUSCAR_ANTIGUEDAD()
        Me.TextBox7.Text = ""
        Me.TextBox8.Text = ""
        Call CARGAR_COMBO_CAT_DE_RESULTADO_DE_EVALUACION()
        Me.TextBox9.Text = ""
        Me.DateTimePicker1.Focus()
    End Sub
    Private Sub BUSCAR_ANTIGUEDAD()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT ID_M_P, FEC_ING_PAME, ANTIGUEDAD_PAME FROM [VISTA MAESTRO DE PERSONAS] WHERE ID_M_P = " & Frm_Ex_Med_Ocup.TextBox36.Text & " order by ID_M_P", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                Me.TextBox5.Text = (MiDataTable.Rows(0).Item(2).ToString) * 12
            Else
                Me.TextBox5.Text = "0"
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        CERRAR = True
        Me.DateTimePicker1.Focus()
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_RESULTADO_DE_EVALUACION()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE RESULTADO DE EVALUACION] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox2
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_R_EVALUACION"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_EVALUACION()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE EVALUACION] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_T_EVALUACION"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub DateTimePicker1_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker1.KeyDown
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
    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
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
    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
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
    Private Sub GENERAR_CODIGO()
        CODIGO = 0
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            MiDataAdapter = New SqlDataAdapter("SELECT ID_EMO FROM [MAESTRO DE HIGIENE EXAMEN MEDICO OCUPACIONAL] ORDER BY ID_EMO", Conexion.CxRRHH)
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
            MsgBox("Debe Digitar el No. de Expediente", vbInformation, "Mensaje del Sistema")
            Me.TextBox2.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Tipo de Evaluación Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox3.Text) = 0 Then
            MsgBox("Debe Digitar el Cargo Anterior Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox3.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox4.Text) = 0 Then
            MsgBox("Debe Digitar el Cargo Actual Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox4.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox6.Text) = 0 Then
            MsgBox("Debe Digitar los Años de Exposición Anterior Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox6.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox5.Text) = 0 Then
            MsgBox("Debe Digitar los Años de Exposición Actual Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox5.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox7.Text) = 0 Then
            MsgBox("Debe Digitar el Diagnostico Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox7.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox8.Text) = 0 Then
            MsgBox("Debe Digitar las Recomendaciones Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox8.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox2.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Resultado de la Evaluacion Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox9.Text) = 0 Then
            MsgBox("Debe Digitar las Observaciones Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox9.Focus()
            Exit Sub
        End If

        Call GRABAR_REGISTRO()
        h5ID_DOS = Me.TextBox1.Text
        CADENAsql = "Select * from [dbo].[VISTA MAESTRO DE HIGIENE EXAMEN MEDICO OCUPACIONAL] WHERE ID_M_P = " & Frm_Ex_Med_Ocup.TextBox36.Text & " ORDER BY FECHA_EVA"
        Call Me.CARGAR_DATAGRIDVIEW()
        Call Me.ARMAR_DATAGRIDVIEW()
        Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        emoID_EMO = Me.TextBox1.Text
        Me.TextBox1.Text = "0"
        Me.TextBox2.Text = ""
        'Call CARGAR_COMBO_CAT_DE_TIPO_DE_EVALUACION()
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox6.Text = ""
        Me.TextBox5.Text = ""
        Me.TextBox7.Text = ""
        Me.TextBox8.Text = ""
        'Call CARGAR_COMBO_CAT_DE_RESULTADO_DE_EVALUACION()
        Me.TextBox9.Text = ""
        Me.DateTimePicker1.Focus()
    End Sub
    Private Sub GRABAR_REGISTRO()
        Dim fFECHA_1 As String
        fFECHA_1 = Mid(Me.DateTimePicker1.Value, 1, 10)
        Dim f1 As Object = If(fFECHA_1 <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_1 + "', 105), 23)", "NULL")

        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE HIGIENE EXAMEN MEDICO OCUPACIONAL] (ID_EMO, DEPARTAMENTO, SERVICIO, CARGO, ID_M_P, FECHA_EVA, NO_EXPEDIENTE, ID_T_EVALUACION, CARGO_ANTERIOR, CARGO_ACTUAL, AE_ANTERIOR, AE_ACTUAL, DIAGNOSTICO, RECOMENDACIONES, ID_R_EVALUACION, OBSERVACIONES)" &
            "values (" & CInt(Me.TextBox1.Text) & ", '" & Frm_Ex_Med_Ocup.TextBox6.Text & "', '" & Frm_Ex_Med_Ocup.TextBox7.Text & "', '" & Frm_Ex_Med_Ocup.TextBox4.Text & "', " & Frm_Ex_Med_Ocup.TextBox36.Text & ", " & f1 & ", '" & Me.TextBox2.Text & "', " & Me.ComboBox1.SelectedValue & ", '" & Me.TextBox3.Text & "', '" & Me.TextBox4.Text & "', '" & Me.TextBox6.Text & "', '" & Me.TextBox5.Text & "', '" & Me.TextBox7.Text & "', '" & Me.TextBox8.Text & "', " & Me.ComboBox2.SelectedValue & ", '" & Me.TextBox9.Text & "')"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW()
        Frm_Ex_Med_Ocup.DGV.Columns(0).HeaderText = "Id"
        Frm_Ex_Med_Ocup.DGV.Columns(0).Width = 30
        Frm_Ex_Med_Ocup.DGV.Columns(0).Visible = True
        Frm_Ex_Med_Ocup.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Ex_Med_Ocup.DGV.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Frm_Ex_Med_Ocup.DGV.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Frm_Ex_Med_Ocup.DGV.Columns(1).HeaderText = "Departamento"
        Frm_Ex_Med_Ocup.DGV.Columns(1).Width = 170
        Frm_Ex_Med_Ocup.DGV.Columns(1).Visible = True
        Frm_Ex_Med_Ocup.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Ex_Med_Ocup.DGV.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Frm_Ex_Med_Ocup.DGV.Columns(2).HeaderText = "Servicio/ Area"
        Frm_Ex_Med_Ocup.DGV.Columns(2).Width = 170
        Frm_Ex_Med_Ocup.DGV.Columns(2).Visible = True
        Frm_Ex_Med_Ocup.DGV.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Ex_Med_Ocup.DGV.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Frm_Ex_Med_Ocup.DGV.Columns(3).HeaderText = "Cargo"
        Frm_Ex_Med_Ocup.DGV.Columns(3).Width = 100
        Frm_Ex_Med_Ocup.DGV.Columns(3).Visible = True
        Frm_Ex_Med_Ocup.DGV.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Ex_Med_Ocup.DGV.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Frm_Ex_Med_Ocup.DGV.Columns(4).HeaderText = "ID_M_P"
        Frm_Ex_Med_Ocup.DGV.Columns(4).Width = 0
        Frm_Ex_Med_Ocup.DGV.Columns(4).Visible = False

        Frm_Ex_Med_Ocup.DGV.Columns(5).HeaderText = "Fecha Evaluacion"
        Frm_Ex_Med_Ocup.DGV.Columns(5).Width = 80
        Frm_Ex_Med_Ocup.DGV.Columns(5).Visible = True
        Frm_Ex_Med_Ocup.DGV.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Ex_Med_Ocup.DGV.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Frm_Ex_Med_Ocup.DGV.Columns(6).HeaderText = "No. Expediente"
        Frm_Ex_Med_Ocup.DGV.Columns(6).Width = 70
        Frm_Ex_Med_Ocup.DGV.Columns(6).Visible = True
        Frm_Ex_Med_Ocup.DGV.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Ex_Med_Ocup.DGV.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Frm_Ex_Med_Ocup.DGV.Columns(7).HeaderText = "ID_T_EVALUACION"
        Frm_Ex_Med_Ocup.DGV.Columns(7).Width = 0
        Frm_Ex_Med_Ocup.DGV.Columns(7).Visible = False

        Frm_Ex_Med_Ocup.DGV.Columns(8).HeaderText = "Tipo Evaluacion"
        Frm_Ex_Med_Ocup.DGV.Columns(8).Width = 90
        Frm_Ex_Med_Ocup.DGV.Columns(8).Visible = True
        Frm_Ex_Med_Ocup.DGV.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Ex_Med_Ocup.DGV.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Frm_Ex_Med_Ocup.DGV.Columns(9).HeaderText = "Cargo Anterior"
        Frm_Ex_Med_Ocup.DGV.Columns(9).Width = 90
        Frm_Ex_Med_Ocup.DGV.Columns(9).Visible = True
        Frm_Ex_Med_Ocup.DGV.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Ex_Med_Ocup.DGV.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Frm_Ex_Med_Ocup.DGV.Columns(10).HeaderText = "Cargo Actual"
        Frm_Ex_Med_Ocup.DGV.Columns(10).Width = 90
        Frm_Ex_Med_Ocup.DGV.Columns(10).Visible = True
        Frm_Ex_Med_Ocup.DGV.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Ex_Med_Ocup.DGV.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Frm_Ex_Med_Ocup.DGV.Columns(11).HeaderText = "Exposicion Anterior"
        Frm_Ex_Med_Ocup.DGV.Columns(11).Width = 90
        Frm_Ex_Med_Ocup.DGV.Columns(11).Visible = True
        Frm_Ex_Med_Ocup.DGV.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Ex_Med_Ocup.DGV.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Frm_Ex_Med_Ocup.DGV.Columns(12).HeaderText = "Exposicion Actual"
        Frm_Ex_Med_Ocup.DGV.Columns(12).Width = 90
        Frm_Ex_Med_Ocup.DGV.Columns(12).Visible = True
        Frm_Ex_Med_Ocup.DGV.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Ex_Med_Ocup.DGV.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Frm_Ex_Med_Ocup.DGV.Columns(13).HeaderText = "Diagnostico"
        Frm_Ex_Med_Ocup.DGV.Columns(13).Width = 0
        Frm_Ex_Med_Ocup.DGV.Columns(13).Visible = False
        Frm_Ex_Med_Ocup.DGV.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Ex_Med_Ocup.DGV.Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Frm_Ex_Med_Ocup.DGV.Columns(14).HeaderText = "Recomendaciones"
        Frm_Ex_Med_Ocup.DGV.Columns(14).Width = 0
        Frm_Ex_Med_Ocup.DGV.Columns(14).Visible = False
        Frm_Ex_Med_Ocup.DGV.Columns(14).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Ex_Med_Ocup.DGV.Columns(14).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Frm_Ex_Med_Ocup.DGV.Columns(15).HeaderText = "IR_R_EVALUACION"
        Frm_Ex_Med_Ocup.DGV.Columns(15).Width = 0
        Frm_Ex_Med_Ocup.DGV.Columns(15).Visible = False
        Frm_Ex_Med_Ocup.DGV.Columns(15).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Ex_Med_Ocup.DGV.Columns(15).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Frm_Ex_Med_Ocup.DGV.Columns(16).HeaderText = "Resultado Evaluacion"
        Frm_Ex_Med_Ocup.DGV.Columns(16).Width = 0
        Frm_Ex_Med_Ocup.DGV.Columns(16).Visible = False
        Frm_Ex_Med_Ocup.DGV.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Ex_Med_Ocup.DGV.Columns(16).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Frm_Ex_Med_Ocup.DGV.Columns(17).HeaderText = "OBSERVACIONES"
        Frm_Ex_Med_Ocup.DGV.Columns(17).Width = 0
        Frm_Ex_Med_Ocup.DGV.Columns(17).Visible = False
        Frm_Ex_Med_Ocup.DGV.Columns(17).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Ex_Med_Ocup.DGV.Columns(17).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

    End Sub
    Private Sub CARGAR_DATAGRIDVIEW()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE HIGIENE EXAMEN MEDICO OCUPACIONAL]")
        Frm_Ex_Med_Ocup.DGV.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE HIGIENE EXAMEN MEDICO OCUPACIONAL]")
        Frm_Ex_Med_Ocup.Label2.Text = Frm_Ex_Med_Ocup.DGV.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        Dim I As Integer
        For I = 0 To Frm_Ex_Med_Ocup.DGV.RowCount - 1
            If Frm_Ex_Med_Ocup.DGV.Item(0, I).Value = emoID_EMO Then
                Frm_Ex_Med_Ocup.DGV.Rows(I).Selected = True
                Frm_Ex_Med_Ocup.DGV.CurrentCell = Frm_Ex_Med_Ocup.DGV.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub

End Class