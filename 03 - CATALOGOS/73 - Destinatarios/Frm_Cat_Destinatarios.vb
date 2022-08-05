Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Cat_Destinatarios
    Private Sub Frm_Cat_Destinatarios_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Cat_Destinatarios_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
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
        Dim MiAdaptador As New SqlClient.SqlDataAdapter("Select * from [VISTA CAT DE DESTINATARIOS] ORDER BY ID_DESTINATARIO", Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA CAT DE DESTINATARIOS]")
        DGV.DataSource = MiDataSet.Tables("[VISTA CAT DE DESTINATARIOS]")
        Me.Label2.Text = "Total de Registros: " & DGV.RowCount
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

        Me.DGV.Columns(2).HeaderText = "Empleado"
        Me.DGV.Columns(2).Width = 200
        Me.DGV.Columns(2).Visible = True
        Me.DGV.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(2).DefaultCellStyle.BackColor = Color.LightGreen

        Me.DGV.Columns(3).HeaderText = "Departamento"
        Me.DGV.Columns(3).Width = 150
        Me.DGV.Columns(3).Visible = True
        Me.DGV.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(4).HeaderText = "Servicio"
        Me.DGV.Columns(4).Width = 150
        Me.DGV.Columns(4).Visible = True
        Me.DGV.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(5).HeaderText = "Cargo"
        Me.DGV.Columns(5).Width = 100
        Me.DGV.Columns(5).Visible = True
        Me.DGV.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(6).HeaderText = "Grado"
        Me.DGV.Columns(6).Width = 100
        Me.DGV.Columns(6).Visible = True
        Me.DGV.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(7).HeaderText = "Correo"
        Me.DGV.Columns(7).Width = 240
        Me.DGV.Columns(7).Visible = True
        Me.DGV.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(8).HeaderText = "Enviar"
        Me.DGV.Columns(8).Width = 60
        Me.DGV.Columns(8).Visible = True
        Me.DGV.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        Dim I As Integer
        For I = 0 To Me.DGV.RowCount - 1
            If Me.DGV.Item(0, I).Value = c44ID_DESTINATARIO Then
                Me.DGV.Rows(I).Selected = True
                Me.DGV.CurrentCell = Me.DGV.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub DGV_Click(sender As Object, e As EventArgs) Handles DGV.Click
        If Me.DGV.RowCount > 0 Then
            For I = 0 To Me.DGV.RowCount - 1
                If DGV.Rows(I).Selected = True Then
                    c44ID_DESTINATARIO = Me.DGV.Item(0, I).Value
                    c44ID_M_P = Me.DGV.Item(1, I).Value
                    c44EMPLEADO = Me.DGV.Item(2, I).Value
                    c44DEPARTAMENTO = Me.DGV.Item(3, I).Value
                    c44SERVICIO = Me.DGV.Item(4, I).Value
                    c44CARGO = Me.DGV.Item(5, I).Value
                    c44GRADO = Me.DGV.Item(6, I).Value
                    c44CORREO = Me.DGV.Item(7, I).Value
                    c44ENVIAR = Me.DGV.Item(8, I).Value
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If c44ID_DESTINATARIO = 0 Then
            MsgBox("Debe seleccionar el registro que desea Eliminar", vbInformation, "Mensaje del Sistema")
            Me.DGV.Focus()
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Esta seguro de Eliminar el registro seleccionado?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            Call ELIMINAR_REGISTRO()
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
        Else
            c44ID_DESTINATARIO = 0
        End If
    End Sub
    Private Sub ELIMINAR_REGISTRO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "DELETE FROM [CAT DE DESTINATARIOS] WHERE ID_DESTINATARIO  = " & CInt(c44ID_DESTINATARIO) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Frm_Cat_Destinatarios_Add.ShowDialog()
        If CERRAR = False Then
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
            c44ID_DESTINATARIO = 0
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If c44ID_DESTINATARIO = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Me.DGV.Focus()
            Exit Sub
        End If
        Frm_Cat_Destinatarios_Upd.ShowDialog()
        If CERRAR = False Then
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
            c44ID_DESTINATARIO = 0
        End If
    End Sub
    Private Sub DGV_DoubleClick(sender As Object, e As EventArgs) Handles DGV.DoubleClick
        If c44ID_DESTINATARIO = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Me.DGV.Focus()
            Exit Sub
        End If
        Frm_Cat_Destinatarios_Upd.ShowDialog()
        If CERRAR = False Then
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
            c44ID_DESTINATARIO = 0
        End If
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            TITULO_EXCEL = "CATALOGO"
            EXPORTAR_DATOS_A_EXCEL(DGV, "DESTINATARIOS (" & UCase(Now) & ")")
        Catch ex As Exception
            MessageBox.Show("No se puede generar el documento Excel.", "Mensaje del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class