Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Situacion_Personal_Proy
    Private Sub Frm_Situacion_Personal_Proy_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Situacion_Personal_Proy_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Me.Close()
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter("Select * from [VISTA MAESTRO DE PROYECCION] where ID_M_P = " & id_EMPLEADO & " ORDER BY FECHA_PROY", Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE PROYECCION]")
        DGV.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE PROYECCION]")
        'Me.Label2.Text = "Total de Registros: " & DGV.RowCount
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

        Me.DGV.Columns(1).HeaderText = "ID_M_P"
        Me.DGV.Columns(1).Width = 0
        Me.DGV.Columns(1).Visible = False

        Me.DGV.Columns(2).HeaderText = "ID_SIT_P"
        Me.DGV.Columns(2).Width = 0
        Me.DGV.Columns(2).Visible = False

        Me.DGV.Columns(3).HeaderText = "Situacion"
        Me.DGV.Columns(3).Width = 150
        Me.DGV.Columns(3).Visible = True
        Me.DGV.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DGV.Columns(3).DefaultCellStyle.BackColor = Color.LightBlue

        Me.DGV.Columns(4).HeaderText = "Fecha Proyeccion"
        Me.DGV.Columns(4).Width = 70
        Me.DGV.Columns(4).Visible = True
        Me.DGV.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(4).DefaultCellStyle.BackColor = Color.Khaki

        Me.DGV.Columns(5).HeaderText = "Observaciones"
        Me.DGV.Columns(5).Width = 370
        Me.DGV.Columns(5).Visible = True
        Me.DGV.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        Dim I As Integer
        For I = 0 To Me.DGV.RowCount - 1
            If Me.DGV.Item(0, I).Value = proyID_PROY Then
                Me.DGV.Rows(I).Selected = True
                Me.DGV.CurrentCell = Me.DGV.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Dim MENSAJE_1
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If proyID_PROY = 0 Then
            MsgBox("Debe seleccionar el registro que desea Eliminar", vbInformation, "Mensaje del Sistema")
            Me.DGV.Focus()
            Exit Sub
        End If
        Call Clases.BUSCAR_FECHA_Y_HORA()
        If CDate(proyFECHA_PROY) < FECHA_SERVIDOR Then
            MENSAJE_1 = MsgBox("La Fecha que intenta eliminar es menor o igual al dia actual, ¿Esta seguro de continuar?", vbInformation + vbYesNo, "Mensaje del Sistema")
            If MENSAJE_1 = vbYes Then
                MENSAJE = MsgBox("¿Esta seguro de Eliminar el registro seleccionado?", vbInformation + vbYesNo, "Mensaje del Sistema")
                If MENSAJE = vbYes Then
                    Call ELIMINAR_REGISTRO()
                    Call CARGAR_DATAGRIDVIEW()
                    Call ARMAR_DATAGRIDVIEW()
                Else
                    proyID_PROY = 0
                End If
            End If
        Else
            MENSAJE = MsgBox("¿Esta seguro de Eliminar el registro seleccionado?", vbInformation + vbYesNo, "Mensaje del Sistema")
            If MENSAJE = vbYes Then
                Call ELIMINAR_REGISTRO()
                Call CARGAR_DATAGRIDVIEW()
                Call ARMAR_DATAGRIDVIEW()
            Else
                proyID_PROY = 0
            End If
        End If
    End Sub
    Private Sub ELIMINAR_REGISTRO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE PROYECCION] WHERE ID_PROY = " & CInt(proyID_PROY) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub DGV_Click(sender As Object, e As EventArgs) Handles DGV.Click
        If Me.DGV.RowCount > 0 Then
            For I = 0 To Me.DGV.RowCount - 1
                If DGV.Rows(I).Selected = True Then
                    proyID_PROY = Me.DGV.Item(0, I).Value
                    proyID_M_P = Me.DGV.Item(1, I).Value
                    proyID_SIT_P = Me.DGV.Item(2, I).Value
                    proyD_SIT_P = Me.DGV.Item(3, I).Value
                    proyFECHA_PROY = Me.DGV.Item(4, I).Value
                    proyOBSERVACIONES = Me.DGV.Item(5, I).Value
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Frm_Situacion_Personal_Proy_Add.ShowDialog()
        If CERRAR = False Then
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
            proyID_PROY = 0
        End If
    End Sub
End Class