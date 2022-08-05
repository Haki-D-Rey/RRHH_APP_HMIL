Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Cat_Estructura_N8
    Private Sub Frm_Cat_Estructura_N8_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Cat_Estructura_N8_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        Dim MiAdaptador As New SqlClient.SqlDataAdapter("SELECT * FROM [dbo].[VISTA CAT DE ESTRUCTURA NIVEL 8] ORDER BY [ORD8]", Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA CAT DE ESTRUCTURA NIVEL 8]")
        DGV.DataSource = MiDataSet.Tables("[VISTA CAT DE ESTRUCTURA NIVEL 8]")
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
        Me.DGV.Columns(2).Width = 100
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
        Me.DGV.Columns(5).Width = 100
        Me.DGV.Columns(5).Visible = True
        Me.DGV.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(6).HeaderText = "ID_N3"
        Me.DGV.Columns(6).Width = 0
        Me.DGV.Columns(6).Visible = False

        Me.DGV.Columns(7).HeaderText = "ORDN3"
        Me.DGV.Columns(7).Width = 0
        Me.DGV.Columns(7).Visible = False

        Me.DGV.Columns(8).HeaderText = "Descripcion N3"
        Me.DGV.Columns(8).Width = 100
        Me.DGV.Columns(8).Visible = True
        Me.DGV.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(9).HeaderText = "ID_N4"
        Me.DGV.Columns(9).Width = 0
        Me.DGV.Columns(9).Visible = False

        Me.DGV.Columns(10).HeaderText = "ORDN4"
        Me.DGV.Columns(10).Width = 0
        Me.DGV.Columns(10).Visible = False

        Me.DGV.Columns(11).HeaderText = "Descripcion N4"
        Me.DGV.Columns(11).Width = 100
        Me.DGV.Columns(11).Visible = True
        Me.DGV.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(12).HeaderText = "ID_N5"
        Me.DGV.Columns(12).Width = 0
        Me.DGV.Columns(12).Visible = False

        Me.DGV.Columns(13).HeaderText = "ORDN5"
        Me.DGV.Columns(13).Width = 0
        Me.DGV.Columns(13).Visible = False

        Me.DGV.Columns(14).HeaderText = "Descripcion N5"
        Me.DGV.Columns(14).Width = 100
        Me.DGV.Columns(14).Visible = True
        Me.DGV.Columns(14).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(15).HeaderText = "ID_N6"
        Me.DGV.Columns(15).Width = 0
        Me.DGV.Columns(15).Visible = False

        Me.DGV.Columns(16).HeaderText = "ORDN6"
        Me.DGV.Columns(16).Width = 0
        Me.DGV.Columns(16).Visible = False

        Me.DGV.Columns(17).HeaderText = "Descripcion N6"
        Me.DGV.Columns(17).Width = 100
        Me.DGV.Columns(17).Visible = True
        Me.DGV.Columns(17).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(18).HeaderText = "ID_N7"
        Me.DGV.Columns(18).Width = 0
        Me.DGV.Columns(18).Visible = False

        Me.DGV.Columns(19).HeaderText = "ORDN7"
        Me.DGV.Columns(19).Width = 0
        Me.DGV.Columns(19).Visible = False

        Me.DGV.Columns(20).HeaderText = "Descripcion N7"
        Me.DGV.Columns(20).Width = 100
        Me.DGV.Columns(20).Visible = True
        Me.DGV.Columns(20).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(21).HeaderText = "Id"
        Me.DGV.Columns(21).Width = 10
        Me.DGV.Columns(21).Visible = True
        Me.DGV.Columns(21).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(21).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV.Columns(21).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(22).HeaderText = "Orden"
        Me.DGV.Columns(22).Width = 40
        Me.DGV.Columns(22).Visible = True
        Me.DGV.Columns(22).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(22).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV.Columns(22).DefaultCellStyle.BackColor = Color.Khaki

        Me.DGV.Columns(23).HeaderText = "Descripcion N8"
        Me.DGV.Columns(23).Width = 100
        Me.DGV.Columns(23).Visible = True
        Me.DGV.Columns(23).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(23).DefaultCellStyle.BackColor = Color.LightGreen

        Me.DGV.Columns(24).HeaderText = "Siglas"
        Me.DGV.Columns(24).Width = 40
        Me.DGV.Columns(24).Visible = True
        Me.DGV.Columns(24).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(25).HeaderText = "ID_T_ES"
        Me.DGV.Columns(25).Width = 0
        Me.DGV.Columns(25).Visible = False
        Me.DGV.Columns(25).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(26).HeaderText = "Tipo"
        Me.DGV.Columns(26).Width = 60
        Me.DGV.Columns(26).Visible = True
        Me.DGV.Columns(26).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(27).HeaderText = "Mixta"
        Me.DGV.Columns(27).Width = 40
        Me.DGV.Columns(27).Visible = True
        Me.DGV.Columns(27).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(28).HeaderText = "Militar"
        Me.DGV.Columns(28).Width = 40
        Me.DGV.Columns(28).Visible = True
        Me.DGV.Columns(28).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(29).HeaderText = "Pame"
        Me.DGV.Columns(29).Width = 40
        Me.DGV.Columns(29).Visible = True
        Me.DGV.Columns(29).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        Dim I As Integer
        For I = 0 To Me.DGV.RowCount - 1
            If Me.DGV.Item(21, I).Value = n8ID_N8 Then
                Me.DGV.Rows(I).Selected = True
                Me.DGV.CurrentCell = Me.DGV.Item(21, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If n8ID_N8 = 0 Then
            MsgBox("Debe seleccionar el registro que desea Eliminar", vbInformation, "Mensaje del Sistema")
            Me.DGV.Focus()
            Exit Sub
        End If
        Call VERIFICAR_ESTRUCTURAS()
        If HAY_CARGOS = True Then
            MsgBox("No se puede Eliminar este registro ya que contiene Cargos con Dependencias", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Esta seguro de Eliminar el registro seleccionado?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            Call ELIMINAR_REGISTRO()
            Call REENUMERAR()
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
        Else
            n8ID_N8 = 0
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
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [MAESTRO DE CARGOS] WHERE ID_N1 = " & n8ID_N1 & " AND ID_N2 = " & n8ID_N2 & " AND ID_N3 = " & n8ID_N3 & " AND ID_N4 = " & n8ID_N4 & " AND ID_N5 = " & n8ID_N5 & " AND ID_N6 = " & n8ID_N6 & " AND ID_N7 = " & n8ID_N7 & " AND ID_N8 = " & n8ID_N8 & "", Conexion.CxRRHH)
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
            MiSqlCommand.CommandText = "DELETE FROM [CAT DE ESTRUCTURA NIVEL 8] WHERE ID_N8 = " & CInt(n8ID_N8) & ""
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
                    n8ID_N1 = Me.DGV.Item(0, I).Value
                    n8ORD1 = Me.DGV.Item(1, I).Value
                    n8D_N1 = Me.DGV.Item(2, I).Value
                    n8ID_N2 = Me.DGV.Item(3, I).Value
                    n8ORD2 = Me.DGV.Item(4, I).Value
                    n8D_N2 = Me.DGV.Item(5, I).Value
                    n8ID_N3 = Me.DGV.Item(6, I).Value
                    n8ORD3 = Me.DGV.Item(7, I).Value
                    n8D_N3 = Me.DGV.Item(8, I).Value
                    n8ID_N4 = Me.DGV.Item(9, I).Value
                    n8ORD4 = Me.DGV.Item(10, I).Value
                    n8D_N4 = Me.DGV.Item(11, I).Value
                    n8ID_N5 = Me.DGV.Item(12, I).Value
                    n8ORD5 = Me.DGV.Item(13, I).Value
                    n8D_N5 = Me.DGV.Item(14, I).Value
                    n8ID_N6 = Me.DGV.Item(15, I).Value
                    n8ORD6 = Me.DGV.Item(16, I).Value
                    n8D_N6 = Me.DGV.Item(17, I).Value
                    n8ID_N7 = Me.DGV.Item(18, I).Value
                    n8ORD7 = Me.DGV.Item(19, I).Value
                    n8D_N7 = Me.DGV.Item(20, I).Value
                    n8ID_N8 = Me.DGV.Item(21, I).Value
                    n8ORD8 = Me.DGV.Item(22, I).Value
                    n8D_N8 = Me.DGV.Item(23, I).Value
                    n8SIGLAS = Me.DGV.Item(24, I).Value
                    n8ID_T_ES = Me.DGV.Item(25, I).Value
                    n8D_T_ES = Me.DGV.Item(26, I).Value
                    n8MIXTA = Me.DGV.Item(27, I).Value
                    n8MILITAR = Me.DGV.Item(28, I).Value
                    n8PAME = Me.DGV.Item(29, I).Value
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Frm_Cat_Estructura_N8_Add.ShowDialog()
        If CERRAR = False Then
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
            n8ID_N8 = 0
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If n8ID_N8 = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Me.DGV.Focus()
            Exit Sub
        End If
        Frm_Cat_Estructura_N8_Upd.ShowDialog()
        If CERRAR = False Then
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
            n8ID_N8 = 0
        End If
    End Sub
    Private Sub DGV_DoubleClick(sender As Object, e As EventArgs) Handles DGV.DoubleClick
        If n8ID_N8 = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Me.DGV.Focus()
            Exit Sub
        End If
        Frm_Cat_Estructura_N8_Upd.ShowDialog()
        If CERRAR = False Then
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
            n8ID_N8 = 0
        End If
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            TITULO_EXCEL = "CATALOGO"
            EXPORTAR_DATOS_A_EXCEL(DGV, "ESTRUCTURA NIVEL 8 (" & UCase(Now) & ")")
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
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [dbo].[VISTA CAT DE ESTRUCTURA NIVEL 8] ORDER BY [ORD8]", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                For I = 0 To MiDataTable.Rows.Count - 1
                    CONTADOR = CONTADOR + 1
                    iINDICE = MiDataTable.Rows(I).Item(21).ToString
                    iNORDEN = CONTADOR
                    Call ACTUALIZAR_CAT_DE_ESTRUCTURA_NIVEL_8()
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
    Private Sub ACTUALIZAR_CAT_DE_ESTRUCTURA_NIVEL_8()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [CAT DE ESTRUCTURA NIVEL 8] SET ORDEN = '" & iNORDEN & "' WHERE ID_N8 = " & CInt(iINDICE) & ""
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
    '    PARAMETRO = 1   'CATALOGO DE NIVEL 8
    '    Frm_Visor_Reportes.CrystalR.ShowExportButton = True
    '    Frm_Visor_Reportes.CrystalR.ShowPrintButton = True
    '    SELECCION = ""
    '    Frm_Visor_Reportes.ShowDialog()
    'End Sub
End Class