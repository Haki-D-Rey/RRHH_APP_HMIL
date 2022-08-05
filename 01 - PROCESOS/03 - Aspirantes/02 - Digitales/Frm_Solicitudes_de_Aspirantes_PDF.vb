Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Solicitudes_de_Aspirantes_PDF
    Private Sub Frm_Solicitudes_de_Aspirantes_PDF_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Solicitudes_de_Aspirantes_PDF_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.Text = "Aspirantes\ Documentos Digitales de: " & aspAN
        CADENAsql = "SELECT * FROM [VISTA MAESTRO DE DOCUMENTOS DIGITALES ASPIRANTES] WHERE ID_S = " & aspID_S & " ORDER BY ID_S, ID_T_DOC_ASP"
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
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE DOCUMENTOS DIGITALES ASPIRANTES]")
        DGV.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE DOCUMENTOS DIGITALES ASPIRANTES]")
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

        Me.DGV.Columns(1).HeaderText = "ID_S"
        Me.DGV.Columns(1).Width = 0
        Me.DGV.Columns(1).Visible = False

        Me.DGV.Columns(2).HeaderText = "ID_D1"
        Me.DGV.Columns(2).Width = 0
        Me.DGV.Columns(2).Visible = False

        Me.DGV.Columns(3).HeaderText = "DIRECTORIO"
        Me.DGV.Columns(3).Width = 0
        Me.DGV.Columns(3).Visible = False

        Me.DGV.Columns(4).HeaderText = "ID_T_DOC"
        Me.DGV.Columns(4).Width = 0
        Me.DGV.Columns(4).Visible = False

        Me.DGV.Columns(5).HeaderText = "Tipo de Documento"
        Me.DGV.Columns(5).Width = 200
        Me.DGV.Columns(5).Visible = True
        Me.DGV.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(6).HeaderText = "Nombre"
        Me.DGV.Columns(6).Width = 300
        Me.DGV.Columns(6).Visible = True
        Me.DGV.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(7).HeaderText = "Observacion"
        Me.DGV.Columns(7).Width = 290
        Me.DGV.Columns(7).Visible = True
        Me.DGV.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub DGV_Click(sender As Object, e As EventArgs) Handles DGV.Click
        If Me.DGV.RowCount > 0 Then
            For I = 0 To Me.DGV.RowCount - 1
                If DGV.Rows(I).Selected = True Then
                    ddaID_M_DD_A = Me.DGV.Item(0, I).Value
                    ddaID_S = Me.DGV.Item(1, I).Value
                    ddaID_D1 = Me.DGV.Item(2, I).Value
                    ddaDIRECTORIO = Me.DGV.Item(3, I).Value
                    ddaID_T_DOC_ASP = Me.DGV.Item(4, I).Value
                    ddaD_T_DOC_ASP = Me.DGV.Item(5, I).Value
                    ddaNOMBRE = Me.DGV.Item(6, I).Value
                    ddaOBSERVACION = Me.DGV.Item(7, I).Value
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub Button44_Click(sender As Object, e As EventArgs) Handles Button44.Click
        Frm_Solicitudes_de_Aspirantes_PDF_Add.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE DOCUMENTOS DIGITALES ASPIRANTES] WHERE ID_S = " & ddaID_S & " ORDER BY ID_S, ID_T_DOC_ASP"
            Call CARGAR_DATAGRIDVIEW()
            Call CARGAR_DATAGRIDVIEW()
        End If
    End Sub
    Private Sub Button42_Click(sender As Object, e As EventArgs) Handles Button42.Click
        If ddaID_M_DD_A = 0 Then
            MsgBox("Debe seleccionar el registro que desea Eliminar", vbInformation, "Mensaje del Sistema")
            Me.DGV.Focus()
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Esta seguro de Eliminar el registro seleccionado?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Conexion.ABD_RRHH()
            Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
            Try

                Dim FILE As String = ddaDIRECTORIO & "\" & ddaNOMBRE

                My.Computer.FileSystem.DeleteFile(FILE, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin, FileIO.UICancelOption.ThrowException)

                MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE DOCUMENTOS DIGITALES ASPIRANTES] WHERE ID_M_DD_A = " & CInt(ddaID_M_DD_A) & ""
                MiSqlCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
        Else
            ddaID_M_DD_A = 0
        End If
    End Sub
    Private Sub Button51_Click(sender As Object, e As EventArgs) Handles Button51.Click
        If ddaID_M_DD_A = 0 Then
            MsgBox("Debe seleccionar el registro que desea Visualizar", vbInformation, "Mensaje del Sistema")
            Me.DGV.Focus()
            Exit Sub
        End If
        vmdDIRECTORIO = ddaDIRECTORIO
        vmdNOMBRE = ddaNOMBRE
        Frm_Visor_Documentos.ShowDialog()
        vmdDIRECTORIO = ""
        vmdNOMBRE = ""
    End Sub
End Class