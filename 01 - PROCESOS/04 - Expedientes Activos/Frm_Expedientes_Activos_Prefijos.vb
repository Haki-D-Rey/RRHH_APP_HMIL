Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Expedientes_Activos_Prefijos
    Private Sub Frm_Expedientes_Activos_Prefijos_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Expedientes_Activos_Prefijos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = "0"
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW()
        CADENAsql = "Select * from [CAT DE PREFIJOS] ORDER BY DESCRIPCION"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[CAT DE PREFIJOS]")
        DGV.DataSource = MiDataSet.Tables("[CAT DE PREFIJOS]")
        Me.Label2.Text = "Total de Registros: " & DGV.RowCount
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW()
        Me.DGV.Columns(0).HeaderText = "Id"
        Me.DGV.Columns(0).Width = 20
        Me.DGV.Columns(0).Visible = True
        Me.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(1).HeaderText = "Prefijo"
        Me.DGV.Columns(1).Width = 120
        Me.DGV.Columns(1).Visible = True
        Me.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub DGV_Click(sender As Object, e As EventArgs) Handles DGV.Click
        If Me.DGV.RowCount > 0 Then
            For I = 0 To Me.DGV.RowCount - 1
                If DGV.Rows(I).Selected = True Then
                    bPREFIJO = Me.DGV.Item(1, I).Value
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub DGV_DoubleClick(sender As Object, e As EventArgs) Handles DGV.DoubleClick
        If bPREFIJO <> "" Then
            Me.TextBox1.Text = "0"
            CERRAR = False
            Me.Close()
        Else
            MsgBox("Debe seleccionar el registro que desea Visualizar", vbInformation, "Mensaje del Sistema")
        End If
    End Sub
End Class