Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Cat_Estructura_N5
    Private Sub Frm_Cat_Estructura_N5_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Cat_Estructura_N5_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        Dim MiAdaptador As New SqlClient.SqlDataAdapter("SELECT * FROM [dbo].[VISTA CAT DE ESTRUCTURA NIVEL 5] ORDER BY [ORD5]", Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA CAT DE ESTRUCTURA NIVEL 5]")
        DGV.DataSource = MiDataSet.Tables("[VISTA CAT DE ESTRUCTURA NIVEL 5]")
        Me.Label2.Text = "Total de Registros: " & DGV.RowCount
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW()
        Me.DGV.Columns(0).HeaderText = "ID_N1"
        Me.DGV.Columns(0).Width = 0
        Me.DGV.Columns(0).Visible = False

        Me.DGV.Columns(1).HeaderText = "ORD1"
        Me.DGV.Columns(1).Width = 0
        Me.DGV.Columns(1).Visible = False

        Me.DGV.Columns(2).HeaderText = "Descripcion N1"
        Me.DGV.Columns(2).Width = 150
        Me.DGV.Columns(2).Visible = True
        Me.DGV.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        'Me.DGV.Columns(2).Frozen = True

        Me.DGV.Columns(3).HeaderText = "ID_N2"
        Me.DGV.Columns(3).Width = 0
        Me.DGV.Columns(3).Visible = False

        Me.DGV.Columns(4).HeaderText = "ORDN2"
        Me.DGV.Columns(4).Width = 0
        Me.DGV.Columns(4).Visible = False

        Me.DGV.Columns(5).HeaderText = "Descripcion N2"
        Me.DGV.Columns(5).Width = 150
        Me.DGV.Columns(5).Visible = True
        Me.DGV.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(6).HeaderText = "ID_N3"
        Me.DGV.Columns(6).Width = 0
        Me.DGV.Columns(6).Visible = False

        Me.DGV.Columns(7).HeaderText = "ORDN3"
        Me.DGV.Columns(7).Width = 0
        Me.DGV.Columns(7).Visible = False

        Me.DGV.Columns(8).HeaderText = "Descripcion N3"
        Me.DGV.Columns(8).Width = 150
        Me.DGV.Columns(8).Visible = True
        Me.DGV.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(9).HeaderText = "ID_N4"
        Me.DGV.Columns(9).Width = 0
        Me.DGV.Columns(9).Visible = False

        Me.DGV.Columns(10).HeaderText = "ORDN4"
        Me.DGV.Columns(10).Width = 0
        Me.DGV.Columns(10).Visible = False

        Me.DGV.Columns(11).HeaderText = "Descripcion N4"
        Me.DGV.Columns(11).Width = 150
        Me.DGV.Columns(11).Visible = True
        Me.DGV.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(12).HeaderText = "Id"
        Me.DGV.Columns(12).Width = 10
        Me.DGV.Columns(12).Visible = True
        Me.DGV.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV.Columns(12).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(13).HeaderText = "Orden"
        Me.DGV.Columns(13).Width = 40
        Me.DGV.Columns(13).Visible = True
        Me.DGV.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV.Columns(13).DefaultCellStyle.BackColor = Color.Khaki

        Me.DGV.Columns(14).HeaderText = "Descripcion N5"
        Me.DGV.Columns(14).Width = 160
        Me.DGV.Columns(14).Visible = True
        Me.DGV.Columns(14).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(14).DefaultCellStyle.BackColor = Color.LightGreen

        Me.DGV.Columns(15).HeaderText = "Siglas"
        Me.DGV.Columns(15).Width = 60
        Me.DGV.Columns(15).Visible = True
        Me.DGV.Columns(15).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(16).HeaderText = "ID_T_ES"
        Me.DGV.Columns(16).Width = 0
        Me.DGV.Columns(16).Visible = False
        Me.DGV.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(17).HeaderText = "Tipo"
        Me.DGV.Columns(17).Width = 80
        Me.DGV.Columns(17).Visible = True
        Me.DGV.Columns(17).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(18).HeaderText = "Mixta"
        Me.DGV.Columns(18).Width = 40
        Me.DGV.Columns(18).Visible = True
        Me.DGV.Columns(18).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(19).HeaderText = "Militar"
        Me.DGV.Columns(19).Width = 40
        Me.DGV.Columns(19).Visible = True
        Me.DGV.Columns(19).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(20).HeaderText = "Pame"
        Me.DGV.Columns(20).Width = 40
        Me.DGV.Columns(20).Visible = True
        Me.DGV.Columns(20).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        Dim I As Integer
        For I = 0 To Me.DGV.RowCount - 1
            If Me.DGV.Item(12, I).Value = n5ID_N5 Then
                Me.DGV.Rows(I).Selected = True
                Me.DGV.CurrentCell = Me.DGV.Item(12, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If n5ID_N5 = 0 Then
            MsgBox("Debe seleccionar el registro que desea Eliminar", vbInformation, "Mensaje del Sistema")
            Me.DGV.Focus()
            Exit Sub
        End If
        Call VERIFICAR_ESTRUCTURAS()
        If HAY_CARGOS = True Then
            MsgBox("No se puede Eliminar este registro ya que contiene Estructuras\ Cargos de Dependencias en el Nivel 6", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Esta seguro de Eliminar el registro seleccionado?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            Call ELIMINAR_REGISTRO()
            Call REENUMERAR()
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
        Else
            n5ID_N5 = 0
        End If
    End Sub
    Private Sub VERIFICAR_ESTRUCTURAS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [MAESTRO DE CARGOS] WHERE ID_N1 = " & n5ID_N1 & " AND ID_N2 = " & n5ID_N2 & " AND ID_N3 = " & n5ID_N3 & " AND ID_N4 = " & n5ID_N4 & " AND ID_N5 = " & n5ID_N5 & "", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                HAY_CARGOS = True
            Else
                HAY_CARGOS = False
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ELIMINAR_REGISTRO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "DELETE FROM [CAT DE ESTRUCTURA NIVEL 5] WHERE ID_N5 = " & CInt(n5ID_N5) & ""
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
                    n5ID_N1 = Me.DGV.Item(0, I).Value
                    n5ORD1 = Me.DGV.Item(1, I).Value
                    n5D_N1 = Me.DGV.Item(2, I).Value
                    n5ID_N2 = Me.DGV.Item(3, I).Value
                    n5ORD2 = Me.DGV.Item(4, I).Value
                    n5D_N2 = Me.DGV.Item(5, I).Value
                    n5ID_N3 = Me.DGV.Item(6, I).Value
                    n5ORD3 = Me.DGV.Item(7, I).Value
                    n5D_N3 = Me.DGV.Item(8, I).Value
                    n5ID_N4 = Me.DGV.Item(9, I).Value
                    n5ORD4 = Me.DGV.Item(10, I).Value
                    n5D_N4 = Me.DGV.Item(11, I).Value
                    n5ID_N5 = Me.DGV.Item(12, I).Value
                    n5ORD5 = Me.DGV.Item(13, I).Value
                    n5D_N5 = Me.DGV.Item(14, I).Value
                    n5SIGLAS = Me.DGV.Item(15, I).Value
                    n5ID_T_ES = Me.DGV.Item(16, I).Value
                    n5D_T_ES = Me.DGV.Item(17, I).Value
                    n5MIXTA = Me.DGV.Item(18, I).Value
                    n5MILITAR = Me.DGV.Item(19, I).Value
                    n5PAME = Me.DGV.Item(20, I).Value
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Frm_Cat_Estructura_N5_Add.ShowDialog()
        If CERRAR = False Then
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
            n5ID_N5 = 0
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If n5ID_N5 = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Me.DGV.Focus()
            Exit Sub
        End If
        Frm_Cat_Estructura_N5_Upd.ShowDialog()
        If CERRAR = False Then
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
            n5ID_N5 = 0
        End If
    End Sub
    Private Sub DGV_DoubleClick(sender As Object, e As EventArgs) Handles DGV.DoubleClick
        If n5ID_N5 = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Me.DGV.Focus()
            Exit Sub
        End If
        Frm_Cat_Estructura_N5_Upd.ShowDialog()
        If CERRAR = False Then
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
            n5ID_N5 = 0
        End If
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            TITULO_EXCEL = "CATALOGO"
            EXPORTAR_DATOS_A_EXCEL(DGV, "ESTRUCTURA NIVEL 5 (" & UCase(Now) & ")")
        Catch ex As Exception
            MessageBox.Show("No se puede generar el documento Excel.", "Mensaje del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Dim iINDICE As Integer
    Dim iNORDEN As String
    Private Sub REENUMERAR()
        Dim CONTADOR As Integer
        CONTADOR = 0
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [dbo].[VISTA CAT DE ESTRUCTURA NIVEL 5] ORDER BY [ORD5]", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                For I = 0 To MiDataTable.Rows.Count - 1
                    CONTADOR = CONTADOR + 1
                    iINDICE = MiDataTable.Rows(I).Item(12).ToString
                    iNORDEN = CONTADOR
                    Call ACTUALIZAR_CAT_DE_ESTRUCTURA_NIVEL_5()
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
    Private Sub ACTUALIZAR_CAT_DE_ESTRUCTURA_NIVEL_5()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [CAT DE ESTRUCTURA NIVEL 5] SET ORDEN = '" & iNORDEN & "' WHERE ID_N5 = " & CInt(iINDICE) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    'Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
    '    PARAMETRO = 1   'CATALOGO DE NIVEL 4
    '    Frm_Visor_Reportes.CrystalR.ShowExportButton = True
    '    Frm_Visor_Reportes.CrystalR.ShowPrintButton = True
    '    SELECCION = ""
    '    Frm_Visor_Reportes.ShowDialog()
    'End Sub
End Class