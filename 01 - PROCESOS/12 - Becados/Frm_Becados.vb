Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Becados
    Private Sub Frm_Becados_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Frm_Principal.Timer1.Stop()
        BECADOS = True
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
    End Sub
    Private Sub Button7_Click(sender As Object, e As System.EventArgs) Handles Button7.Click    'CERRAR
        BECADOS = False
        Frm_Principal.Timer1.Start()
        Me.Close()
    End Sub
    Private Sub Frm_Becados_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            BECADOS = False
            Frm_Principal.Timer1.Start()
            Me.Close()
        End If
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter("Select * from [VISTA MAESTRO DE BECADOS] ORDER BY NOMBRES, APELLIDOS", Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE BECADOS]")
        DGV.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE BECADOS]")
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

        Me.DGV.Columns(1).HeaderText = "ID_M_P"
        Me.DGV.Columns(1).Width = 0
        Me.DGV.Columns(1).Visible = False
        Me.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(2).HeaderText = "Nombres"
        Me.DGV.Columns(2).Width = 120
        Me.DGV.Columns(2).Visible = True
        Me.DGV.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(2).DefaultCellStyle.BackColor = Color.Khaki

        Me.DGV.Columns(3).HeaderText = "Apellidos"
        Me.DGV.Columns(3).Width = 120
        Me.DGV.Columns(3).Visible = True
        Me.DGV.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(3).DefaultCellStyle.BackColor = Color.Khaki

        Me.DGV.Columns(4).HeaderText = "Codigo"
        Me.DGV.Columns(4).Width = 70
        Me.DGV.Columns(4).Visible = True
        Me.DGV.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(5).HeaderText = "Cedula"
        Me.DGV.Columns(5).Width = 90
        Me.DGV.Columns(5).Visible = True
        Me.DGV.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(6).HeaderText = "No. MINSA"
        Me.DGV.Columns(6).Width = 50
        Me.DGV.Columns(6).Visible = True
        Me.DGV.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        '
        Me.DGV.Columns(7).HeaderText = "D_T_SEXO"
        Me.DGV.Columns(7).Width = 0
        Me.DGV.Columns(7).Visible = False

        Me.DGV.Columns(8).HeaderText = "NOMBRE_ESTUDIO"
        Me.DGV.Columns(8).Width = 0
        Me.DGV.Columns(8).Visible = False

        Me.DGV.Columns(9).HeaderText = "ID_PAIS"
        Me.DGV.Columns(9).Width = 0
        Me.DGV.Columns(9).Visible = False

        Me.DGV.Columns(10).HeaderText = "D_PAIS"
        Me.DGV.Columns(10).Width = 0
        Me.DGV.Columns(10).Visible = False

        Me.DGV.Columns(11).HeaderText = "CENTRO_ESTUDIO"
        Me.DGV.Columns(11).Width = 0
        Me.DGV.Columns(11).Visible = False

        Me.DGV.Columns(12).HeaderText = "Fecha Inicio"
        Me.DGV.Columns(12).Width = 70
        Me.DGV.Columns(12).Visible = True
        Me.DGV.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(13).HeaderText = "Fecha Fin"
        Me.DGV.Columns(13).Width = 70
        Me.DGV.Columns(13).Visible = True
        Me.DGV.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(14).HeaderText = "TELEFONO"
        Me.DGV.Columns(14).Width = 0
        Me.DGV.Columns(14).Visible = False

        Me.DGV.Columns(15).HeaderText = "No. Cuenta"
        Me.DGV.Columns(15).Width = 0
        Me.DGV.Columns(15).Visible = False
        Me.DGV.Columns(15).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(16).HeaderText = "Correo"
        Me.DGV.Columns(16).Width = 0
        Me.DGV.Columns(16).Visible = False
        Me.DGV.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(17).HeaderText = "Matri cula"
        Me.DGV.Columns(17).Width = 60
        Me.DGV.Columns(17).Visible = True
        Me.DGV.Columns(17).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(17).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV.Columns(17).DefaultCellStyle.Format = "##,##0.00"

        Me.DGV.Columns(18).HeaderText = "Inscrip cion"
        Me.DGV.Columns(18).Width = 60
        Me.DGV.Columns(18).Visible = True
        Me.DGV.Columns(18).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(18).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV.Columns(18).DefaultCellStyle.Format = "##,##0.00"

        Me.DGV.Columns(19).HeaderText = "Alimen tacion"
        Me.DGV.Columns(19).Width = 60
        Me.DGV.Columns(19).Visible = True
        Me.DGV.Columns(19).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(19).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV.Columns(19).DefaultCellStyle.Format = "##,##0.00"

        Me.DGV.Columns(20).HeaderText = "Hospe daje"
        Me.DGV.Columns(20).Width = 60
        Me.DGV.Columns(20).Visible = True
        Me.DGV.Columns(20).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(20).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV.Columns(20).DefaultCellStyle.Format = "##,##0.00"

        Me.DGV.Columns(21).HeaderText = "Pago Titulo"
        Me.DGV.Columns(21).Width = 60
        Me.DGV.Columns(21).Visible = True
        Me.DGV.Columns(21).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(21).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV.Columns(21).DefaultCellStyle.Format = "##,##0.00"

        Me.DGV.Columns(22).HeaderText = "Otros"
        Me.DGV.Columns(22).Width = 60
        Me.DGV.Columns(22).Visible = True
        Me.DGV.Columns(22).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(22).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV.Columns(22).DefaultCellStyle.Format = "##,##0.00"

        Me.DGV.Columns(23).HeaderText = "Total"
        Me.DGV.Columns(23).Width = 65
        Me.DGV.Columns(23).Visible = True
        Me.DGV.Columns(23).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(23).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV.Columns(23).DefaultCellStyle.Format = "##,##0.00"

        Me.DGV.Columns(24).HeaderText = "OBSERVACIONES"
        Me.DGV.Columns(24).Width = 0
        Me.DGV.Columns(24).Visible = False
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        Dim I As Integer
        For I = 0 To Me.DGV.RowCount - 1
            If Me.DGV.Item(0, I).Value = mbID_B Then
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
                    Dim FILA = DGV.CurrentRow.Index
                    If Not IsDBNull(Me.DGV.Item(0, FILA).Value) Then : mbID_B = Me.DGV.Item(0, FILA).Value : Else : mbID_B = 0 : End If
                    If Not IsDBNull(Me.DGV.Item(1, FILA).Value) Then : mbID_M_P = Me.DGV.Item(1, FILA).Value : Else : mbID_M_P = 0 : End If
                    If Not IsDBNull(Me.DGV.Item(2, FILA).Value) Then : mbNOMBRES = Me.DGV.Item(2, FILA).Value : Else : mbNOMBRES = "" : End If
                    If Not IsDBNull(Me.DGV.Item(3, FILA).Value) Then : mbAPELLIDOS = Me.DGV.Item(3, FILA).Value : Else : mbAPELLIDOS = "" : End If
                    If Not IsDBNull(Me.DGV.Item(4, FILA).Value) Then : mbCODIGO = Me.DGV.Item(4, FILA).Value : Else : mbCODIGO = "" : End If
                    If Not IsDBNull(Me.DGV.Item(5, FILA).Value) Then : mbN_CEDULA = Me.DGV.Item(5, FILA).Value : Else : mbN_CEDULA = "" : End If
                    If Not IsDBNull(Me.DGV.Item(6, FILA).Value) Then : mbN_MINSA = Me.DGV.Item(6, FILA).Value : Else : mbN_MINSA = "" : End If
                    If Not IsDBNull(Me.DGV.Item(7, FILA).Value) Then : mbD_T_SEXO = Me.DGV.Item(7, FILA).Value : Else : mbD_T_SEXO = "" : End If
                    If Not IsDBNull(Me.DGV.Item(8, FILA).Value) Then : mbNOMBRE_ESTUDIO = Me.DGV.Item(8, FILA).Value : Else : mbNOMBRE_ESTUDIO = "" : End If
                    If Not IsDBNull(Me.DGV.Item(9, FILA).Value) Then : mbID_PAIS = Me.DGV.Item(9, FILA).Value : Else : mbID_PAIS = 0 : End If
                    If Not IsDBNull(Me.DGV.Item(10, FILA).Value) Then : mbD_PAIS = Me.DGV.Item(10, FILA).Value : Else : mbD_PAIS = "" : End If
                    If Not IsDBNull(Me.DGV.Item(11, FILA).Value) Then : mbCENTRO_ESTUDIO = Me.DGV.Item(11, FILA).Value : Else : mbCENTRO_ESTUDIO = "" : End If
                    If Not IsDBNull(Me.DGV.Item(12, FILA).Value) Then : mbFEC_INI = Me.DGV.Item(12, FILA).Value : Else : mbFEC_INI = "" : End If
                    If Not IsDBNull(Me.DGV.Item(13, FILA).Value) Then : mbFEC_FIN = Me.DGV.Item(13, FILA).Value : Else : mbFEC_FIN = "" : End If
                    If Not IsDBNull(Me.DGV.Item(14, FILA).Value) Then : mbTELEFONO = Me.DGV.Item(14, FILA).Value : Else : mbTELEFONO = "" : End If
                    If Not IsDBNull(Me.DGV.Item(15, FILA).Value) Then : mbNO_CUENTA = Me.DGV.Item(15, FILA).Value : Else : mbNO_CUENTA = "" : End If
                    If Not IsDBNull(Me.DGV.Item(16, FILA).Value) Then : mbCORREO = Me.DGV.Item(16, FILA).Value : Else : mbCORREO = "" : End If
                    If Not IsDBNull(Me.DGV.Item(17, FILA).Value) Then : mbMATRICULA = Me.DGV.Item(17, FILA).Value : Else : mbMATRICULA = "" : End If
                    If Not IsDBNull(Me.DGV.Item(18, FILA).Value) Then : mbINSCRIPCION = Me.DGV.Item(18, FILA).Value : Else : mbINSCRIPCION = "" : End If
                    If Not IsDBNull(Me.DGV.Item(19, FILA).Value) Then : mbALIMENTACION = Me.DGV.Item(19, FILA).Value : Else : mbALIMENTACION = "" : End If
                    If Not IsDBNull(Me.DGV.Item(20, FILA).Value) Then : mbHOSPEDAJE = Me.DGV.Item(20, FILA).Value : Else : mbHOSPEDAJE = "" : End If
                    If Not IsDBNull(Me.DGV.Item(21, FILA).Value) Then : mbPAGO_TITULO = Me.DGV.Item(21, FILA).Value : Else : mbPAGO_TITULO = "" : End If
                    If Not IsDBNull(Me.DGV.Item(22, FILA).Value) Then : mbOTROS = Me.DGV.Item(22, FILA).Value : Else : mbOTROS = "" : End If
                    If Not IsDBNull(Me.DGV.Item(23, FILA).Value) Then : mbTOTAL = Me.DGV.Item(23, FILA).Value : Else : mbTOTAL = "" : End If
                    If Not IsDBNull(Me.DGV.Item(24, FILA).Value) Then : mbOBSERVACIONES = Me.DGV.Item(24, FILA).Value : Else : mbOBSERVACIONES = "" : End If
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click   'ELIMINAR
        If mbID_B = 0 Then
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
            mbID_B = 0
        End If
    End Sub
    Private Sub ELIMINAR_REGISTRO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE BECADOS] WHERE ID_B = " & CInt(mbID_B) & ""
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
        Frm_Becados_Agregar.ShowDialog()
        If CERRAR = False Then
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
            mbID_B = 0
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click   'EDITAR
        If mbID_B = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Me.DGV.Focus()
            Exit Sub
        End If
        Frm_Becados_Editar.ShowDialog()
        If CERRAR = False Then
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
            mbID_B = 0
        End If
    End Sub
    Private Sub DGV_DoubleClick(sender As Object, e As EventArgs) Handles DGV.DoubleClick
        If mbID_B = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Me.DGV.Focus()
            Exit Sub
        End If
        Frm_Becados_Editar.ShowDialog()
        If CERRAR = False Then
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
            mbID_B = 0
        End If
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click   'IMPRIMIR
        Frm_Becados_Imprimir.ShowDialog()
        'If mbID_B = 0 Then
        '    MsgBox("Debe seleccionar el registro del Aspirante que desea Imprimir", vbInformation, "Mensaje del Sistema")
        '    Exit Sub
        'End If
        'PARAMETRO = 36
        'Frm_Visor_Reportes.CrystalReportViewer1.ShowExportButton = True
        'Frm_Visor_Reportes.CrystalReportViewer1.ShowPrintButton = True
        'SELECCION = "{MAESTRO_DE_BECADOS.ID_B} = " & mbID_B
        Frm_Visor_Reportes.ShowDialog()
    End Sub
    Private Sub Frm_Becados_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        BECADOS = False
    End Sub
    Private Sub Frm_Becados_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        BECADOS = False
    End Sub
    Private Sub Frm_Becados_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        BECADOS = False
    End Sub
    Private Sub Frm_Becados_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        BECADOS = False
    End Sub
End Class