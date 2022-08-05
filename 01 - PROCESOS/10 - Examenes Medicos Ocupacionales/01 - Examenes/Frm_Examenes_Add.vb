Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Examenes_Add
    Private Sub Frm_Examenes_Add_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            CERRAR = True
            Me.ComboBox1.Focus()
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Examenes_Add_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = "0"
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_EXAMENES()
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_RESULTADO_DE_EXAMEN()
        Me.TextBox7.Text = ""
        Me.ComboBox1.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        CERRAR = True
        Me.ComboBox1.Focus()
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_EXAMENES()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE EXAMENES] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_T_EXAMENES"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_RESULTADO_DE_EXAMEN()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE RESULTADO DE EXAMEN] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox2
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_TRE"
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
    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
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
            MiDataAdapter = New SqlDataAdapter("SELECT ID_EX FROM [MAESTRO DE HIGIENE EXAMEN MEDICO OCUPACIONAL EXAMENES] ORDER BY ID_EX", Conexion.CxRRHH)
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
            MsgBox("Debe seleccionar el Tipo de Examen Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox2.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Tipo de Resultado de Examen Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox7.Text) = 0 Then
            Me.TextBox7.Text = "."
        End If
        Call GRABAR_REGISTRO()
        exID_EX = Me.TextBox1.Text
        CADENAsql = "Select * from [dbo].[VISTA MAESTRO DE HIGIENE EXAMEN MEDICO OCUPACIONAL EXAMENES] WHERE ID_EMO = " & emoID_EMO & " ORDER BY D_T_EXAMENES"
        Call Me.CARGAR_DATAGRIDVIEW()
        Call Me.ARMAR_DATAGRIDVIEW()
        Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        Me.TextBox1.Text = "0"
        Me.TextBox7.Text = ""
        Me.ComboBox1.Focus()
    End Sub
    Private Sub GRABAR_REGISTRO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE HIGIENE EXAMEN MEDICO OCUPACIONAL EXAMENES] (ID_EX, ID_EMO, ID_T_EXAMENES, ID_TRE, OBSERVACIONES)" &
            "values (" & CInt(Me.TextBox1.Text) & ", " & emoID_EMO & ", " & Me.ComboBox1.SelectedValue & ", " & Me.ComboBox2.SelectedValue & ", '" & Me.TextBox7.Text & "')"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
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
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE HIGIENE EXAMEN MEDICO OCUPACIONAL EXAMENES]")
        Frm_Examenes.DGV.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE HIGIENE EXAMEN MEDICO OCUPACIONAL EXAMENES]")
        Frm_Examenes.Label2.Text = Frm_Examenes.DGV.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW()
        Frm_Examenes.DGV.Columns(0).HeaderText = "Id"
        Frm_Examenes.DGV.Columns(0).Width = 30
        Frm_Examenes.DGV.Columns(0).Visible = True
        Frm_Examenes.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Examenes.DGV.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Frm_Examenes.DGV.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Frm_Examenes.DGV.Columns(1).HeaderText = "ID_DEMO"
        Frm_Examenes.DGV.Columns(1).Width = 0
        Frm_Examenes.DGV.Columns(1).Visible = False
        Frm_Examenes.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Examenes.DGV.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Frm_Examenes.DGV.Columns(2).HeaderText = "ID_T_EXAMENES"
        Frm_Examenes.DGV.Columns(2).Width = 0
        Frm_Examenes.DGV.Columns(2).Visible = False
        Frm_Examenes.DGV.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Examenes.DGV.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Frm_Examenes.DGV.Columns(3).HeaderText = "Examen"
        Frm_Examenes.DGV.Columns(3).Width = 150
        Frm_Examenes.DGV.Columns(3).Visible = True
        Frm_Examenes.DGV.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Examenes.DGV.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Frm_Examenes.DGV.Columns(4).HeaderText = "ID_TRE"
        Frm_Examenes.DGV.Columns(4).Width = 0
        Frm_Examenes.DGV.Columns(4).Visible = False
        Frm_Examenes.DGV.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Examenes.DGV.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Frm_Examenes.DGV.Columns(5).HeaderText = "Tipo de Resultado"
        Frm_Examenes.DGV.Columns(5).Width = 150
        Frm_Examenes.DGV.Columns(5).Visible = True
        Frm_Examenes.DGV.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Examenes.DGV.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Frm_Examenes.DGV.Columns(6).HeaderText = "Observaciones"
        Frm_Examenes.DGV.Columns(6).Width = 290
        Frm_Examenes.DGV.Columns(6).Visible = True
        Frm_Examenes.DGV.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Examenes.DGV.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        Dim I As Integer
        For I = 0 To Frm_Examenes.DGV.RowCount - 1
            If Frm_Examenes.DGV.Item(0, I).Value = exID_EX Then
                Frm_Examenes.DGV.Rows(I).Selected = True
                Frm_Examenes.DGV.CurrentCell = Frm_Examenes.DGV.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
End Class