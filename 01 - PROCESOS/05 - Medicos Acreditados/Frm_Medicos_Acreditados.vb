Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Medicos_Acreditados
    Private Sub Frm_Medicos_Acreditados_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Frm_Principal.Timer1.Stop()
        ACREDITADOS = True
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
    End Sub
    Private Sub Button7_Click(sender As Object, e As System.EventArgs) Handles Button7.Click     'CERRAR
        Frm_Principal.Timer1.Start()
        ACREDITADOS = False
        Me.Close()
    End Sub
    Private Sub Frm_Medicos_Acreditados_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            ACREDITADOS = False
            Frm_Principal.Timer1.Start()
            Me.Close()
        End If
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter("Select * from [VISTA MAESTRO DE MEDICOS ACREDITADOS] ORDER BY APELLIDOS, NOMBRES", Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE MEDICOS ACREDITADOS]")
        DGV.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE MEDICOS ACREDITADOS]")
        Me.Label4.Text = DGV.RowCount & " Registro(s)"
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

        Me.DGV.Columns(1).HeaderText = "Apellidos"
        Me.DGV.Columns(1).Width = 150
        Me.DGV.Columns(1).Visible = True
        Me.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(2).HeaderText = "Nombres"
        Me.DGV.Columns(2).Width = 150
        Me.DGV.Columns(2).Visible = True
        Me.DGV.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(3).HeaderText = "ID_E_MED"
        Me.DGV.Columns(3).Width = 0
        Me.DGV.Columns(3).Visible = False

        Me.DGV.Columns(4).HeaderText = "Especialidad Medica"
        Me.DGV.Columns(4).Width = 170
        Me.DGV.Columns(4).Visible = True
        Me.DGV.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(4).Frozen = True

        Me.DGV.Columns(5).HeaderText = "Cedula"
        Me.DGV.Columns(5).Width = 100
        Me.DGV.Columns(5).Visible = True
        Me.DGV.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(5).DefaultCellStyle.BackColor = Color.Khaki

        Me.DGV.Columns(6).HeaderText = "No. MINSA"
        Me.DGV.Columns(6).Width = 50
        Me.DGV.Columns(6).Visible = True
        Me.DGV.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(6).DefaultCellStyle.BackColor = Color.Khaki

        Me.DGV.Columns(7).HeaderText = "Telefono 1"
        Me.DGV.Columns(7).Width = 80
        Me.DGV.Columns(7).Visible = True
        Me.DGV.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(8).HeaderText = "Telefono 2"
        Me.DGV.Columns(8).Width = 80
        Me.DGV.Columns(8).Visible = True
        Me.DGV.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(9).HeaderText = "CORREO_1"
        Me.DGV.Columns(9).Width = 0
        Me.DGV.Columns(9).Visible = False

        Me.DGV.Columns(10).HeaderText = "CORREO_2"
        Me.DGV.Columns(10).Width = 0
        Me.DGV.Columns(10).Visible = False

        Me.DGV.Columns(11).HeaderText = "ID_NACIONALIDAD_D"
        Me.DGV.Columns(11).Width = 0
        Me.DGV.Columns(11).Visible = False

        Me.DGV.Columns(12).HeaderText = "Nacionalidad"
        Me.DGV.Columns(12).Width = 100
        Me.DGV.Columns(12).Visible = True
        Me.DGV.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(13).HeaderText = "Direccion"
        Me.DGV.Columns(13).Width = 190
        Me.DGV.Columns(13).Visible = True
        Me.DGV.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(14).HeaderText = "ID_T_SEXO"
        Me.DGV.Columns(14).Width = 0
        Me.DGV.Columns(14).Visible = False

        Me.DGV.Columns(15).HeaderText = "Sexo"
        Me.DGV.Columns(15).Width = 70
        Me.DGV.Columns(15).Visible = True
        Me.DGV.Columns(15).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(16).HeaderText = "Fecha Nacimiento"
        Me.DGV.Columns(16).Width = 70
        Me.DGV.Columns(16).Visible = True
        Me.DGV.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(17).HeaderText = "Fecha Desde"
        Me.DGV.Columns(17).Width = 70
        Me.DGV.Columns(17).Visible = True
        Me.DGV.Columns(17).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(18).HeaderText = "Fecha Hasta"
        Me.DGV.Columns(18).Width = 70
        Me.DGV.Columns(18).Visible = True
        Me.DGV.Columns(18).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(19).HeaderText = "FACEBOOK"
        Me.DGV.Columns(19).Width = 0
        Me.DGV.Columns(19).Visible = False

        Me.DGV.Columns(20).HeaderText = "TWITTER"
        Me.DGV.Columns(20).Width = 0
        Me.DGV.Columns(20).Visible = False

        Me.DGV.Columns(21).HeaderText = "Whatsapp"
        Me.DGV.Columns(21).Width = 80
        Me.DGV.Columns(21).Visible = True
        Me.DGV.Columns(21).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(22).HeaderText = "Observaciones"
        Me.DGV.Columns(22).Width = 150
        Me.DGV.Columns(22).Visible = True
        Me.DGV.Columns(22).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(23).HeaderText = "AN"
        Me.DGV.Columns(23).Width = 0
        Me.DGV.Columns(23).Visible = False

        Me.DGV.Columns(24).HeaderText = "FOTO"
        Me.DGV.Columns(24).Width = 0
        Me.DGV.Columns(24).Visible = False

    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        Dim I As Integer
        For I = 0 To Me.DGV.RowCount - 1
            If Me.DGV.Item(0, I).Value = vMAc_ID_MA Then
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
                    If Not IsDBNull(Me.DGV.Item(0, FILA).Value) Then : vMAc_ID_MA = Me.DGV.Item(0, FILA).Value : Else : vMAc_ID_MA = 0 : End If
                    If Not IsDBNull(Me.DGV.Item(1, FILA).Value) Then : vMAc_APELLIDOS = Me.DGV.Item(1, FILA).Value : Else : vMAc_APELLIDOS = "" : End If
                    If Not IsDBNull(Me.DGV.Item(2, FILA).Value) Then : vMAc_NOMBRES = Me.DGV.Item(2, FILA).Value : Else : vMAc_NOMBRES = "" : End If
                    If Not IsDBNull(Me.DGV.Item(3, FILA).Value) Then : vMAc_ID_E_MED = Me.DGV.Item(3, FILA).Value : Else : vMAc_ID_E_MED = 0 : End If
                    If Not IsDBNull(Me.DGV.Item(4, FILA).Value) Then : vMAc_D_E_MED = Me.DGV.Item(4, FILA).Value : Else : vMAc_D_E_MED = "" : End If
                    If Not IsDBNull(Me.DGV.Item(5, FILA).Value) Then : vMAc_N_CEDULA = Me.DGV.Item(5, FILA).Value : Else : vMAc_N_CEDULA = "" : End If
                    If Not IsDBNull(Me.DGV.Item(6, FILA).Value) Then : vMAc_N_MINSA = Me.DGV.Item(6, FILA).Value : Else : vMAc_N_MINSA = "" : End If
                    If Not IsDBNull(Me.DGV.Item(7, FILA).Value) Then : vMAc_TELEFONO_1 = Me.DGV.Item(7, FILA).Value : Else : vMAc_TELEFONO_1 = "" : End If
                    If Not IsDBNull(Me.DGV.Item(8, FILA).Value) Then : vMAc_TELEFONO_2 = Me.DGV.Item(8, FILA).Value : Else : vMAc_TELEFONO_2 = "" : End If
                    If Not IsDBNull(Me.DGV.Item(9, FILA).Value) Then : vMAc_CORREO_1 = Me.DGV.Item(9, FILA).Value : Else : vMAc_CORREO_1 = "" : End If
                    If Not IsDBNull(Me.DGV.Item(10, FILA).Value) Then : vMAc_CORREO_2 = Me.DGV.Item(10, FILA).Value : Else : vMAc_CORREO_2 = "" : End If
                    If Not IsDBNull(Me.DGV.Item(11, FILA).Value) Then : vMAc_ID_NACIONALIDAD_D = Me.DGV.Item(11, FILA).Value : Else : vMAc_ID_NACIONALIDAD_D = 0 : End If
                    If Not IsDBNull(Me.DGV.Item(12, FILA).Value) Then : vMAc_D_NACIONALIDAD_D = Me.DGV.Item(12, FILA).Value : Else : vMAc_D_NACIONALIDAD_D = "" : End If
                    If Not IsDBNull(Me.DGV.Item(13, FILA).Value) Then : vMAc_DIRECCION_D = Me.DGV.Item(13, FILA).Value : Else : vMAc_DIRECCION_D = "" : End If
                    If Not IsDBNull(Me.DGV.Item(14, FILA).Value) Then : vMAc_ID_T_SEXO = Me.DGV.Item(14, FILA).Value : Else : vMAc_ID_T_SEXO = 0 : End If
                    If Not IsDBNull(Me.DGV.Item(15, FILA).Value) Then : vMAc_D_T_SEXO = Me.DGV.Item(15, FILA).Value : Else : vMAc_D_T_SEXO = "" : End If
                    If Not IsDBNull(Me.DGV.Item(16, FILA).Value) Then : vMAc_FEC_NAC = Me.DGV.Item(16, FILA).Value : Else : vMAc_FEC_NAC = "" : End If
                    If Not IsDBNull(Me.DGV.Item(17, FILA).Value) Then : vMAc_FEC_DESDE = Me.DGV.Item(17, FILA).Value : Else : vMAc_FEC_DESDE = "" : End If
                    If Not IsDBNull(Me.DGV.Item(18, FILA).Value) Then : vMAc_FEC_HASTA = Me.DGV.Item(18, FILA).Value : Else : vMAc_FEC_HASTA = "" : End If
                    If Not IsDBNull(Me.DGV.Item(19, FILA).Value) Then : vMAc_FACEBOOK = Me.DGV.Item(19, FILA).Value : Else : vMAc_FACEBOOK = "" : End If
                    If Not IsDBNull(Me.DGV.Item(20, FILA).Value) Then : vMAc_TWITTER = Me.DGV.Item(20, FILA).Value : Else : vMAc_TWITTER = "" : End If
                    If Not IsDBNull(Me.DGV.Item(21, FILA).Value) Then : vMAc_WHATSAPP = Me.DGV.Item(21, FILA).Value : Else : vMAc_WHATSAPP = "" : End If
                    If Not IsDBNull(Me.DGV.Item(22, FILA).Value) Then : vMAc_OBSERVACIONES = Me.DGV.Item(22, FILA).Value : Else : vMAc_OBSERVACIONES = "" : End If
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click    'ELIMINAR
        Call VERIFICAR_ACCESOS("156") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vMAc_ID_MA = 0 Then
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
            vMAc_ID_MA = 0
        End If
    End Sub
    Private Sub ELIMINAR_REGISTRO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE MEDICOS ACREDITADOS] WHERE ID_MA = " & CInt(vMAc_ID_MA) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click    'NUEVO
        Call VERIFICAR_ACCESOS("154") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Medicos_Acreditados_Agregar.ShowDialog()
        If CERRAR = False Then
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
            vMAc_ID_MA = 0
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click    'EDITAR
        Call VERIFICAR_ACCESOS("155") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vMAc_ID_MA = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Me.DGV.Focus()
            Exit Sub
        End If
        Frm_Medicos_Acreditados_Editar.ShowDialog()
        If CERRAR = False Then
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
            vMAc_ID_MA = 0
        End If
    End Sub
    Private Sub DGV_DoubleClick(sender As Object, e As EventArgs) Handles DGV.DoubleClick
        Call VERIFICAR_ACCESOS("155") : If HAY_ACCESO = False Then : Exit Sub : End If
        If vMAc_ID_MA = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Me.DGV.Focus()
            Exit Sub
        End If
        Frm_Medicos_Acreditados_Editar.ShowDialog()
        If CERRAR = False Then
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
            vMAc_ID_MA = 0
        End If
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click    'IMPRIMIR
        Call VERIFICAR_ACCESOS("157") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Medicos_Acreditados_Imprimir.ShowDialog()
    End Sub
    Private Sub Frm_Medicos_Acreditados_Closing(sender As Object, e As CancelEventArgs) Handles MyBase.Closing
        ACREDITADOS = False
    End Sub
    Private Sub Frm_Medicos_Acreditados_Closed(sender As Object, e As EventArgs) Handles MyBase.Closed
        ACREDITADOS = False
    End Sub
    Private Sub Frm_Medicos_Acreditados_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        ACREDITADOS = False
    End Sub
    Private Sub Frm_Medicos_Acreditados_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ACREDITADOS = False
    End Sub
    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click  'BUSCAR
        Call VERIFICAR_ACCESOS("153") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Medicos_Acreditados_Buscar.ShowDialog()
        If CERRAR = False Then
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        End If
    End Sub
End Class