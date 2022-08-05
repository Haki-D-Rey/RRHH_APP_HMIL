Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Public Class Frm_Examenes
    Private Sub Frm_Examenes_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Examenes_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        CADENAsql = "Select * from [dbo].[VISTA MAESTRO DE HIGIENE EXAMEN MEDICO OCUPACIONAL EXAMENES] WHERE ID_EMO = " & emoID_EMO & " ORDER BY D_T_EXAMENES"
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
    End Sub
    Private Sub Button7_Click(sender As Object, e As System.EventArgs) Handles Button7.Click
        Me.Close()
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE HIGIENE EXAMEN MEDICO OCUPACIONAL EXAMENES]")
        DGV.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE HIGIENE EXAMEN MEDICO OCUPACIONAL EXAMENES]")
        Me.Label2.Text = DGV.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW()
        Me.DGV.Columns(0).HeaderText = "Id"
        Me.DGV.Columns(0).Width = 30
        Me.DGV.Columns(0).Visible = True
        Me.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(1).HeaderText = "ID_DEMO"
        Me.DGV.Columns(1).Width = 0
        Me.DGV.Columns(1).Visible = False
        Me.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(2).HeaderText = "ID_T_EXAMENES"
        Me.DGV.Columns(2).Width = 0
        Me.DGV.Columns(2).Visible = False
        Me.DGV.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(3).HeaderText = "Examen"
        Me.DGV.Columns(3).Width = 150
        Me.DGV.Columns(3).Visible = True
        Me.DGV.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(4).HeaderText = "ID_TRE"
        Me.DGV.Columns(4).Width = 0
        Me.DGV.Columns(4).Visible = False
        Me.DGV.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(5).HeaderText = "Tipo de Resultado"
        Me.DGV.Columns(5).Width = 150
        Me.DGV.Columns(5).Visible = True
        Me.DGV.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(6).HeaderText = "Observaciones"
        Me.DGV.Columns(6).Width = 290
        Me.DGV.Columns(6).Visible = True
        Me.DGV.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
    End Sub
    Private Sub DGV_Click(sender As Object, e As EventArgs) Handles DGV.Click
        If Me.DGV.RowCount > 0 Then
            For I = 0 To Me.DGV.RowCount - 1
                If DGV.Rows(I).Selected = True Then
                    exID_EX = Me.DGV.Item(0, I).Value
                    exID_EMO = Me.DGV.Item(1, I).Value
                    exID_T_EXAMENES = Me.DGV.Item(2, I).Value
                    exD_T_EXAMENES = Me.DGV.Item(3, I).Value
                    exID_TRE = Me.DGV.Item(4, I).Value
                    exD_TRE = Me.DGV.Item(5, I).Value
                    exOBSERVACIONES = Me.DGV.Item(6, I).Value
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        Dim I As Integer
        For I = 0 To Me.DGV.RowCount - 1
            If Me.DGV.Item(0, I).Value = exID_EX Then
                Me.DGV.Rows(I).Selected = True
                Me.DGV.CurrentCell = Me.DGV.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Frm_Examenes_Add.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "Select * from [dbo].[VISTA MAESTRO DE HIGIENE EXAMEN MEDICO OCUPACIONAL EXAMENES] WHERE ID_EMO = " & emoID_EMO & " ORDER BY D_T_EXAMENES"
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        End If
    End Sub
    Dim ID As Integer
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Me.CheckBox1.Checked = False Then
            If Val(exID_EX) = 0 Then
                MsgBox("Debe seleccionar el registro que desea Eliminar", vbInformation, "Mensaje del Sistema")
                Exit Sub
            End If
            MENSAJE = MsgBox("¿Esta seguro de Eliminar el registro seleccionado?", vbInformation + vbYesNo, "Mensaje del Sistema")
            If MENSAJE = vbYes Then
                Call BORRAR_REGISTRO_2()
                CADENAsql = "Select * from [dbo].[VISTA MAESTRO DE HIGIENE EXAMEN MEDICO OCUPACIONAL EXAMENES] WHERE ID_EMO = " & emoID_EMO & " ORDER BY D_T_EXAMENES"
                Call CARGAR_DATAGRIDVIEW()
                Call ARMAR_DATAGRIDVIEW()
                exID_EX = 0
            End If
        End If
        If Me.CheckBox1.Checked = True Then
            MENSAJE = MsgBox("¿Esta seguro de Eliminar el (los) registro (s) seleccionado (s)?", vbInformation + vbYesNo, "Mensaje del Sistema")
            If MENSAJE = vbYes Then
                Me.Cursor = Cursors.WaitCursor
                For F = 0 To DGV.RowCount - 1
                    Application.DoEvents()
                    'If DGV.Rows(F).Selected = True Then
                    If DGV.Rows(F).Selected = True Then
                        'DGV.CurrentCell = DGV.Item(0, F)
                        ID = DGV.Item(0, F).Value
                        exID_EX = ID
                        Call BORRAR_REGISTRO_2()
                    End If
                Next
                CADENAsql = "Select * from [dbo].[VISTA MAESTRO DE HIGIENE EXAMEN MEDICO OCUPACIONAL EXAMENES] WHERE ID_EMO = " & emoID_EMO & " ORDER BY D_T_EXAMENES"
                Call CARGAR_DATAGRIDVIEW()
                Call ARMAR_DATAGRIDVIEW()
                exID_EX = 0
                Me.Cursor = Cursors.Default
            End If
        End If
    End Sub
    Private Sub BORRAR_REGISTRO_2()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE HIGIENE EXAMEN MEDICO OCUPACIONAL EXAMENES] WHERE ID_EX = " & CInt(exID_EX) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Val(exID_EX) = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Examenes_Upd.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "Select * from [dbo].[VISTA MAESTRO DE HIGIENE EXAMEN MEDICO OCUPACIONAL EXAMENES] WHERE ID_EMO = " & emoID_EMO & " ORDER BY D_T_EXAMENES"
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        End If
    End Sub
    Private Sub DGV_DoubleClick(sender As Object, e As EventArgs) Handles DGV.DoubleClick
        If Val(exID_EX) = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Examenes_Upd.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "Select * from [dbo].[VISTA MAESTRO DE HIGIENE EXAMEN MEDICO OCUPACIONAL EXAMENES] WHERE ID_EMO = " & emoID_EMO & " ORDER BY D_T_EXAMENES"
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        End If
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        MENSAJE = MsgBox("¿Esta seguro de Generar Todos los Examenes para este Empleado?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            Call RECORRER_CAT_DE_TIPO_DE_EXAMENES()
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
        End If
    End Sub
    Dim TIPO_EXAMEN As Integer
    Private Sub RECORRER_CAT_DE_TIPO_DE_EXAMENES()
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE EXAMENES] ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                For I = 0 To MiDataTable.Rows.Count - 1
                    Call GENERAR_CODIGO()
                    TIPO_EXAMEN = MiDataTable.Rows(I).Item(0).ToString
                    Call VERIFICAR_SI_YA_EXISTE_EXAMEN()
                    If YA_EXISTE_EXAMEN = False Then
                        Call AGREGAR_REGISTRO()
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub AGREGAR_REGISTRO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE HIGIENE EXAMEN MEDICO OCUPACIONAL EXAMENES] (ID_EX, ID_EMO, ID_T_EXAMENES, ID_TRE, OBSERVACIONES)" &
            "values (" & CInt(CODIGO_NUEVO) & ", " & emoID_EMO & ", " & TIPO_EXAMEN & ", '1', '.')"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim YA_EXISTE_EXAMEN As Boolean
    Private Sub VERIFICAR_SI_YA_EXISTE_EXAMEN()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [MAESTRO DE HIGIENE EXAMEN MEDICO OCUPACIONAL EXAMENES] WHERE ID_EMO = " & emoID_EMO & " AND ID_T_EXAMENES = " & TIPO_EXAMEN & "", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                YA_EXISTE_EXAMEN = True
            Else
                YA_EXISTE_EXAMEN = False
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub

    Dim CODIGO_NUEVO As Integer
    Private Sub GENERAR_CODIGO()
        CODIGO_NUEVO = 0
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
                For I1 = 0 To MiDataTable.Rows.Count - 1
                    CODIGO = MiDataTable.Rows(I1).Item(0).ToString
                Next
                CODIGO += 1
            Else
                CODIGO = 1
            End If
            CODIGO_NUEVO = CODIGO
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Me.DGV.MultiSelect = Me.CheckBox1.Checked
    End Sub
End Class