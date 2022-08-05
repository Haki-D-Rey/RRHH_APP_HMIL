Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Public Class Frm_Personal_Tercerizado
    Private Sub Frm_Personal_Tercerizado_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            TERCERIZADOS = False
            Frm_Principal.Timer1.Start()
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Personal_Tercerizado_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Frm_Principal.Timer1.Stop()
        TERCERIZADOS = True
        RemoveHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
        Call CARGAR_COMBO_CAT_DE_EMPRESAS_TERCERIZADAS()
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
        AddHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_EMPRESAS_TERCERIZADAS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE EMPRESAS TERCERIZADAS] ORDER BY ID_E_T", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_E_T"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        TERCERIZADOS = False
        Frm_Principal.Timer1.Start()
        Me.Close()
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter("Select ID_P_T, ID_E_T, D_E_T, CODIGO, APELLIDOS, NOMBRES, AN, N_CEDULA, ID_T_SEXO, D_T_SEXO, TELEFONO_1, TELEFONO_2, D_CARGO_ES, ID_ESTADO_P, D_ESTADO_P, DIRECCION, NSS, OBSERVACIONES, F_INGRESO, F_BAJA, CAUSA_RETIRO from [VISTA MAESTRO DE PERSONAL TERCERIZADO] WHERE ID_E_T = " & Me.ComboBox1.SelectedValue & " ORDER BY CODIGO", Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE PERSONAL TERCERIZADO]")
        DGV.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE PERSONAL TERCERIZADO]")
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

        Me.DGV.Columns(1).HeaderText = "ID_E_T"
        Me.DGV.Columns(1).Width = 0
        Me.DGV.Columns(1).Visible = False

        Me.DGV.Columns(2).HeaderText = "Empresa"
        Me.DGV.Columns(2).Width = 0
        Me.DGV.Columns(2).Visible = False
        Me.DGV.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(3).HeaderText = "Codigo"
        Me.DGV.Columns(3).Width = 80
        Me.DGV.Columns(3).Visible = True
        Me.DGV.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(4).HeaderText = "APELLIDOS"
        Me.DGV.Columns(4).Width = 0
        Me.DGV.Columns(4).Visible = False

        Me.DGV.Columns(5).HeaderText = "NOMBRES"
        Me.DGV.Columns(5).Width = 0
        Me.DGV.Columns(5).Visible = False

        Me.DGV.Columns(6).HeaderText = "Apellidos y Nombres"
        Me.DGV.Columns(6).Width = 230
        Me.DGV.Columns(6).Visible = True
        Me.DGV.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(7).HeaderText = "Cedula"
        Me.DGV.Columns(7).Width = 110
        Me.DGV.Columns(7).Visible = True
        Me.DGV.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(8).HeaderText = "ID_T_SEXO"
        Me.DGV.Columns(8).Width = 0
        Me.DGV.Columns(8).Visible = False

        Me.DGV.Columns(9).HeaderText = "Sexo"
        Me.DGV.Columns(9).Width = 70
        Me.DGV.Columns(9).Visible = True
        Me.DGV.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(9).Frozen = True

        Me.DGV.Columns(10).HeaderText = "Telefono (Claro)"
        Me.DGV.Columns(10).Width = 80
        Me.DGV.Columns(10).Visible = True
        Me.DGV.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(11).HeaderText = "Telefono (Tigo)"
        Me.DGV.Columns(11).Width = 80
        Me.DGV.Columns(11).Visible = True
        Me.DGV.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(12).HeaderText = "Cargo"
        Me.DGV.Columns(12).Width = 200
        Me.DGV.Columns(12).Visible = True
        Me.DGV.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(13).HeaderText = "ID_ESTADO_P"
        Me.DGV.Columns(13).Width = 0
        Me.DGV.Columns(13).Visible = False

        Me.DGV.Columns(14).HeaderText = "Estado"
        Me.DGV.Columns(14).Width = 70
        Me.DGV.Columns(14).Visible = True
        Me.DGV.Columns(14).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(15).HeaderText = "Direccion"
        Me.DGV.Columns(15).Width = 250
        Me.DGV.Columns(15).Visible = True
        Me.DGV.Columns(15).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(16).HeaderText = "Nss"
        Me.DGV.Columns(16).Width = 50
        Me.DGV.Columns(16).Visible = True
        Me.DGV.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(17).HeaderText = "Observaciones"
        Me.DGV.Columns(17).Width = 250
        Me.DGV.Columns(17).Visible = True
        Me.DGV.Columns(17).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(18).HeaderText = "Fecha Ingreso"
        Me.DGV.Columns(18).Width = 80
        Me.DGV.Columns(18).Visible = True
        Me.DGV.Columns(18).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(19).HeaderText = "Fecha Baja"
        Me.DGV.Columns(19).Width = 80
        Me.DGV.Columns(19).Visible = True
        Me.DGV.Columns(19).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(20).HeaderText = "Causa de Retiro"
        Me.DGV.Columns(20).Width = 150
        Me.DGV.Columns(20).Visible = True
        Me.DGV.Columns(20).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        Dim I As Integer
        For I = 0 To Me.DGV.RowCount - 1
            If Me.DGV.Item(0, I).Value = ptID_P_T Then
                Me.DGV.Rows(I).Selected = True
                Me.DGV.CurrentCell = Me.DGV.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub DGV_RowPrePaint(sender As Object, e As DataGridViewRowPrePaintEventArgs) Handles DGV.RowPrePaint
        Call FORMATO_FILAS_DATAGRIDVIEW()
    End Sub
    Private Sub FORMATO_FILAS_DATAGRIDVIEW()
        Dim style As New DataGridViewCellStyle
        style.Font = New Font(DGV.Font, FontStyle.Bold)
        For Each Row As DataGridViewRow In DGV.Rows
            If Row.Cells(14).Value = "NO ACTIVO" Then
                Row.DefaultCellStyle.ForeColor = Color.Red
            End If
        Next
    End Sub
    Private Sub DGV_Click(sender As Object, e As EventArgs) Handles DGV.Click
        If Me.DGV.RowCount > 0 Then
            For I = 0 To Me.DGV.RowCount - 1
                If DGV.Rows(I).Selected = True Then
                    Dim FILA = DGV.CurrentRow.Index
                    If Not IsDBNull(Me.DGV.Item(0, FILA).Value) Then : ptID_P_T = Me.DGV.Item(0, FILA).Value : Else : ptID_P_T = 0 : End If
                    If Not IsDBNull(Me.DGV.Item(1, FILA).Value) Then : ptID_E_T = Me.DGV.Item(1, FILA).Value : Else : ptID_E_T = 0 : End If
                    If Not IsDBNull(Me.DGV.Item(2, FILA).Value) Then : ptD_E_T = Me.DGV.Item(2, FILA).Value : Else : ptD_E_T = "" : End If
                    If Not IsDBNull(Me.DGV.Item(3, FILA).Value) Then : ptCODIGO = Me.DGV.Item(3, FILA).Value : Else : ptCODIGO = "" : End If
                    If Not IsDBNull(Me.DGV.Item(4, FILA).Value) Then : ptAPELLIDOS = Me.DGV.Item(4, FILA).Value : Else : ptAPELLIDOS = "" : End If
                    If Not IsDBNull(Me.DGV.Item(5, FILA).Value) Then : ptNOMBRES = Me.DGV.Item(5, FILA).Value : Else : ptNOMBRES = "" : End If
                    If Not IsDBNull(Me.DGV.Item(6, FILA).Value) Then : ptAN = Me.DGV.Item(6, FILA).Value : Else : ptAN = "" : End If
                    If Not IsDBNull(Me.DGV.Item(7, FILA).Value) Then : ptN_CEDULA = Me.DGV.Item(7, FILA).Value : Else : ptN_CEDULA = "" : End If
                    If Not IsDBNull(Me.DGV.Item(8, FILA).Value) Then : ptID_T_SEXO = Me.DGV.Item(8, FILA).Value : Else : ptID_T_SEXO = 0 : End If
                    If Not IsDBNull(Me.DGV.Item(9, FILA).Value) Then : ptD_T_SEXO = Me.DGV.Item(9, FILA).Value : Else : ptD_T_SEXO = "" : End If
                    If Not IsDBNull(Me.DGV.Item(10, FILA).Value) Then : ptTELEFONO_1 = Me.DGV.Item(10, FILA).Value : Else : vMAc_CORREO_2 = "" : End If
                    If Not IsDBNull(Me.DGV.Item(11, FILA).Value) Then : ptTELEFONO_2 = Me.DGV.Item(11, FILA).Value : Else : ptTELEFONO_2 = "" : End If
                    If Not IsDBNull(Me.DGV.Item(12, FILA).Value) Then : ptD_CARGO_ES = Me.DGV.Item(12, FILA).Value : Else : ptD_CARGO_ES = "" : End If
                    If Not IsDBNull(Me.DGV.Item(13, FILA).Value) Then : ptID_ESTADO_P = Me.DGV.Item(13, FILA).Value : Else : ptID_ESTADO_P = 0 : End If
                    If Not IsDBNull(Me.DGV.Item(14, FILA).Value) Then : ptD_ESTADO_P = Me.DGV.Item(14, FILA).Value : Else : ptD_ESTADO_P = 0 : End If
                    If Not IsDBNull(Me.DGV.Item(15, FILA).Value) Then : ptDIRECCION = Me.DGV.Item(15, FILA).Value : Else : ptDIRECCION = "" : End If
                    If Not IsDBNull(Me.DGV.Item(16, FILA).Value) Then : ptNSS = Me.DGV.Item(16, FILA).Value : Else : ptNSS = "" : End If
                    If Not IsDBNull(Me.DGV.Item(17, FILA).Value) Then : ptOBSERVACIONES = Me.DGV.Item(17, FILA).Value : Else : ptOBSERVACIONES = "" : End If
                    If Not IsDBNull(Me.DGV.Item(18, FILA).Value) Then : ptF_INGRESO = Me.DGV.Item(18, FILA).Value : Else : ptF_INGRESO = "" : End If
                    If Not IsDBNull(Me.DGV.Item(19, FILA).Value) Then : ptF_BAJA = Me.DGV.Item(19, FILA).Value : Else : ptF_BAJA = "" : End If
                    If Not IsDBNull(Me.DGV.Item(20, FILA).Value) Then : ptCAUSA_RETIRO = Me.DGV.Item(20, FILA).Value : Else : ptCAUSA_RETIRO = "" : End If
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click   'ELIMINAR
        Call VERIFICAR_ACCESOS("221") : If HAY_ACCESO = False Then : Exit Sub : End If
        If ptID_P_T = 0 Then
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
            ptID_P_T = 0
        End If
    End Sub
    Private Sub ELIMINAR_REGISTRO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE PERSONAL TERCERIZADO] WHERE ID_P_T = " & CInt(ptID_P_T) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click   'NUEVO
        Call VERIFICAR_ACCESOS("219") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Personal_Tercerizado_Add.ShowDialog()
        If CERRAR = False Then
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
            ptID_P_T = 0
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click   'EDITAR
        Call VERIFICAR_ACCESOS("220") : If HAY_ACCESO = False Then : Exit Sub : End If
        If ptID_P_T = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Me.DGV.Focus()
            Exit Sub
        End If
        Frm_Personal_Tercerizado_Upd.ShowDialog()
        If CERRAR = False Then
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
            ptID_P_T = 0
        End If
    End Sub
    Private Sub DGV_DoubleClick(sender As Object, e As EventArgs) Handles DGV.DoubleClick
        Call VERIFICAR_ACCESOS("220") : If HAY_ACCESO = False Then : Exit Sub : End If
        If ptID_P_T = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Me.DGV.Focus()
            Exit Sub
        End If
        Frm_Personal_Tercerizado_Upd.ShowDialog()
        If CERRAR = False Then
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
            ptID_P_T = 0
        End If
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click   'IMPRIMIR
        Call VERIFICAR_ACCESOS("222") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Personal_Tercerizado_Imprimir.ShowDialog()
    End Sub
    Private Sub Frm_Personal_Tercerizado_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        TERCERIZADOS = False
    End Sub
    Private Sub Frm_Personal_Tercerizado_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        TERCERIZADOS = False
    End Sub
    Private Sub Frm_Personal_Tercerizado_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        TERCERIZADOS = False
    End Sub
    Private Sub Frm_Personal_Tercerizado_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        TERCERIZADOS = False
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call VERIFICAR_ACCESOS("218") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Personal_Tercerizado_Buscar.ShowDialog()
        If CERRAR = False Then
            Me.ComboBox1.Text = ptD_E_T
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
            ptID_P_T = 0
        End If
    End Sub
End Class