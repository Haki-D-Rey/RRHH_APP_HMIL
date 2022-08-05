Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop.Excel.ApplicationClass
Imports Microsoft.Office.Interop
Public Class Frm_Documentos
    Private Sub Frm_Documentos_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Documentos_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.Text = "Documentos Recepcionados de: " & aspAN
        Call VERIFICAR_SI_EXISTEN_REGISTROS()
        If EXISTEN_REGISTROS = False Then
            Call RECORRER_DOCUMENTOS()
        End If
        CADENAsql = "SELECT * FROM [VISTA SOL_MAESTRO DE SOLICITANTES DOCUMENTOS] WHERE ID_S = " & aspID_S & " ORDER BY ID_T_DOC_ASP"
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA SOL_MAESTRO DE SOLICITANTES DOCUMENTOS]")
        DGV.DataSource = MiDataSet.Tables("[VISTA SOL_MAESTRO DE SOLICITANTES DOCUMENTOS]")
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW()
        Me.DGV.Columns(0).HeaderText = "Id"
        Me.DGV.Columns(0).Width = 10
        Me.DGV.Columns(0).Visible = True
        Me.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV.Columns(0).DefaultCellStyle.BackColor = Color.Black
        Me.DGV.Columns(0).ReadOnly = True

        Me.DGV.Columns(1).HeaderText = "ID_S"
        Me.DGV.Columns(1).Width = 0
        Me.DGV.Columns(1).Visible = False
        Me.DGV.Columns(1).ReadOnly = True

        Me.DGV.Columns(2).HeaderText = "ID_T_DOC_ASP"
        Me.DGV.Columns(2).Width = 0
        Me.DGV.Columns(2).Visible = False
        Me.DGV.Columns(2).ReadOnly = True

        Me.DGV.Columns(3).HeaderText = "Tipo de Documento"
        Me.DGV.Columns(3).Width = 250
        Me.DGV.Columns(3).Visible = True
        Me.DGV.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DGV.Columns(3).DefaultCellStyle.BackColor = Color.Khaki
        Me.DGV.Columns(3).ReadOnly = True

        Me.DGV.Columns(4).HeaderText = "Recepcionado"
        Me.DGV.Columns(4).Width = 80
        Me.DGV.Columns(4).Visible = True
        Me.DGV.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(5).HeaderText = "Observaciones"
        Me.DGV.Columns(5).Width = 390
        Me.DGV.Columns(5).Visible = True
        Me.DGV.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
    End Sub
    Dim aID_T_DOC_ASP As Integer
    Private Sub RECORRER_DOCUMENTOS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE DOCUMENTO ASPIRANTES] ORDER BY ID_T_DOC_ASP", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                For I1 = 0 To MiDataTable.Rows.Count - 1
                    aID_T_DOC_ASP = MiDataTable.Rows(I1).Item(0).ToString
                    Call GENERAR_CODIGO()
                    Call GRABAR()
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
    Private Sub GRABAR()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [SOL_MAESTRO DE SOLICITANTES DOCUMENTOS] (ID_D_R, ID_S, ID_T_DOC_ASP, RECEPCIONADO, OBSERVACIONES) " &
                                       "values (" & C & ", " & aspID_S & ", " & aID_T_DOC_ASP & ", 'TRUE', '')"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim C As Integer
    Private Sub GENERAR_CODIGO()
        CODIGO = 0
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            MiDataAdapter = New SqlDataAdapter("SELECT ID_D_R FROM [SOL_MAESTRO DE SOLICITANTES DOCUMENTOS] ORDER BY ID_D_R", Conexion.CxRRHH)
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
            C = CODIGO
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim EXISTEN_REGISTROS As Boolean
    Private Sub VERIFICAR_SI_EXISTEN_REGISTROS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [SOL_MAESTRO DE SOLICITANTES DOCUMENTOS] WHERE ID_S = " & aspID_S & "", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                EXISTEN_REGISTROS = True
            Else
                EXISTEN_REGISTROS = False
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim value1 As Boolean
    Dim value2 As String
    Private Sub DGV_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DGV.CellEndEdit
        If (e.ColumnIndex = 4) Then
            value1 = DGV.CurrentCell.Value.ToString
            Dim row As DataGridViewRow = DGV.CurrentRow
            Call ACTUALIZAR_CELDAS_1()
        End If
        If (e.ColumnIndex = 5) Then
            value2 = DGV.CurrentCell.Value.ToString
            Dim row As DataGridViewRow = DGV.CurrentRow
            Call ACTUALIZAR_CELDAS_2()
        Else
            Return
        End If
    End Sub
    Private Sub ACTUALIZAR_CELDAS_1()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [SOL_MAESTRO DE SOLICITANTES DOCUMENTOS] SET RECEPCIONADO = '" & value1 & "' WHERE ID_D_R = " & CInt(dID_D_R) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ACTUALIZAR_CELDAS_2()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [SOL_MAESTRO DE SOLICITANTES DOCUMENTOS] SET OBSERVACIONES = '" & Trim(UCase(value2)) & "' WHERE ID_D_R = " & CInt(dID_D_R) & ""
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
        If DGV.SelectedCells.Count <> 0 Then
            Dim FILA = DGV.CurrentRow.Index
            If Not IsDBNull(Me.DGV.Item(0, FILA).Value) Then : dID_D_R = Me.DGV.Item(0, FILA).Value : Else : dID_D_R = 0 : End If
            If Not IsDBNull(Me.DGV.Item(1, FILA).Value) Then : dID_S = Me.DGV.Item(1, FILA).Value : Else : dID_S = 0 : End If
            If Not IsDBNull(Me.DGV.Item(2, FILA).Value) Then : dID_T_DOC_ASP = Me.DGV.Item(2, FILA).Value : Else : dID_T_DOC_ASP = 0 : End If
            If Not IsDBNull(Me.DGV.Item(3, FILA).Value) Then : dD_T_DOC_ASP = Me.DGV.Item(3, FILA).Value : Else : dD_T_DOC_ASP = "" : End If
            If Not IsDBNull(Me.DGV.Item(4, FILA).Value) Then : dRECEPCIONADO = Me.DGV.Item(4, FILA).Value : Else : dRECEPCIONADO = False : End If
            If Not IsDBNull(Me.DGV.Item(5, FILA).Value) Then : dOBSERVACIONES = Me.DGV.Item(5, FILA).Value : Else : dOBSERVACIONES = "" : End If
        End If
    End Sub
    Private Sub DGV_SelectionChanged(sender As Object, e As EventArgs) Handles DGV.SelectionChanged
        If DGV.SelectedCells.Count <> 0 Then
            Dim FILA = DGV.CurrentRow.Index
            If Not IsDBNull(Me.DGV.Item(0, FILA).Value) Then : dID_D_R = Me.DGV.Item(0, FILA).Value : Else : dID_D_R = 0 : End If
            If Not IsDBNull(Me.DGV.Item(1, FILA).Value) Then : dID_S = Me.DGV.Item(1, FILA).Value : Else : dID_S = 0 : End If
            If Not IsDBNull(Me.DGV.Item(2, FILA).Value) Then : dID_T_DOC_ASP = Me.DGV.Item(2, FILA).Value : Else : dID_T_DOC_ASP = 0 : End If
            If Not IsDBNull(Me.DGV.Item(3, FILA).Value) Then : dD_T_DOC_ASP = Me.DGV.Item(3, FILA).Value : Else : dD_T_DOC_ASP = "" : End If
            If Not IsDBNull(Me.DGV.Item(4, FILA).Value) Then : dRECEPCIONADO = Me.DGV.Item(4, FILA).Value : Else : dRECEPCIONADO = False : End If
            If Not IsDBNull(Me.DGV.Item(5, FILA).Value) Then : dOBSERVACIONES = Me.DGV.Item(5, FILA).Value : Else : dOBSERVACIONES = "" : End If
        End If
    End Sub
End Class