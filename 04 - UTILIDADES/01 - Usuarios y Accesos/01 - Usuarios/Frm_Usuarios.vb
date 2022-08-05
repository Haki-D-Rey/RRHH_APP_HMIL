Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Usuarios
    Private Sub Frm_Usuarios_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.ComboBox2.Focus()
            uID_USUARIO = 0
            USUARIOS = False
            Frm_Principal.Timer1.Start()
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Usuarios_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Frm_Principal.Timer1.Stop()
        USUARIOS = True
        RemoveHandler Me.ComboBox2.SelectedIndexChanged, AddressOf ComboBox2_SelectedIndexChanged
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_USUARIO()
        CADENAsql = "Select * from [VISTA MAESTRO DE USUARIOS] WHERE (ID_TIPO_USUARIO = " & Me.ComboBox2.SelectedValue & ") ORDER BY USUARIO"
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
        AddHandler Me.ComboBox2.SelectedIndexChanged, AddressOf ComboBox2_SelectedIndexChanged
        uID_USUARIO = 0
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Me.ComboBox2.Focus()
        USUARIOS = False
        uID_USUARIO = 0
        Frm_Principal.Timer1.Start()
        Me.Close()
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        CADENAsql = "Select * from [VISTA MAESTRO DE USUARIOS] WHERE (ID_TIPO_USUARIO = " & Me.ComboBox2.Text & ") ORDER BY USUARIO"
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
    End Sub
    Dim CADENAsqlCombo As String
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_USUARIO()
        If isuID_TIPO_USUARIO = 1 Then
            CADENAsqlCombo = "Select * FROM [CAT DE TIPO DE USUARIO] ORDER BY DESCRIPCION"
        Else
            CADENAsqlCombo = "SELECT * FROM [CAT DE TIPO DE USUARIO] WHERE ID_TIPO_USUARIO <> 1 ORDER BY DESCRIPCION"
        End If
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter(CADENAsqlCombo, Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox2
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_TIPO_USUARIO"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter("Select * from [VISTA MAESTRO DE USUARIOS] WHERE ID_TIPO_USUARIO = " & Me.ComboBox2.SelectedValue & " ORDER BY D_TIPO_USUARIO, USUARIO", Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE USUARIOS]")
        DGV.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE USUARIOS]")
        Me.Label2.Text = DGV.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW()
        Me.DGV.Columns(0).HeaderText = "."
        Me.DGV.Columns(0).Width = 10
        Me.DGV.Columns(0).Visible = True
        Me.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(1).HeaderText = "ID_T_USUARIO"
        Me.DGV.Columns(1).Width = 0
        Me.DGV.Columns(1).Visible = False

        Me.DGV.Columns(2).HeaderText = "Tipo"
        Me.DGV.Columns(2).Width = 80
        Me.DGV.Columns(2).Visible = True
        Me.DGV.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(3).HeaderText = "Nombres y Apellidos"
        Me.DGV.Columns(3).Width = 140
        Me.DGV.Columns(3).Visible = True
        Me.DGV.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(3).DefaultCellStyle.BackColor = Color.LightGreen

        Me.DGV.Columns(4).HeaderText = "Cuenta"
        Me.DGV.Columns(4).Width = 100
        Me.DGV.Columns(4).Visible = True
        Me.DGV.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(5).HeaderText = "Contraseña"
        Me.DGV.Columns(5).Width = 0
        Me.DGV.Columns(5).Visible = False

        Me.DGV.Columns(6).HeaderText = "ID_SESION"
        Me.DGV.Columns(6).Width = 0
        Me.DGV.Columns(6).Visible = False

        Me.DGV.Columns(7).HeaderText = "Estado Sesion"
        Me.DGV.Columns(7).Width = 60
        Me.DGV.Columns(7).Visible = True
        Me.DGV.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(8).HeaderText = "ID_ESTADO"
        Me.DGV.Columns(8).Width = 0
        Me.DGV.Columns(8).Visible = False

        Me.DGV.Columns(9).HeaderText = "Estado Usuario"
        Me.DGV.Columns(9).Width = 60
        Me.DGV.Columns(9).Visible = True
        Me.DGV.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(10).HeaderText = "Codigo"
        Me.DGV.Columns(10).Width = 70
        Me.DGV.Columns(10).Visible = True
        Me.DGV.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(11).HeaderText = "Cargo"
        Me.DGV.Columns(11).Width = 70
        Me.DGV.Columns(11).Visible = True
        Me.DGV.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(12).HeaderText = "Nivel 1"
        Me.DGV.Columns(12).Width = 60
        Me.DGV.Columns(12).Visible = True
        Me.DGV.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(13).HeaderText = "Nivel 2"
        Me.DGV.Columns(13).Width = 60
        Me.DGV.Columns(13).Visible = True
        Me.DGV.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(14).HeaderText = "Nivel 3"
        Me.DGV.Columns(14).Width = 60
        Me.DGV.Columns(14).Visible = True
        Me.DGV.Columns(14).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(15).HeaderText = "Nivel 4"
        Me.DGV.Columns(15).Width = 60
        Me.DGV.Columns(15).Visible = True
        Me.DGV.Columns(15).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(16).HeaderText = "Nivel 5"
        Me.DGV.Columns(16).Width = 60
        Me.DGV.Columns(16).Visible = True
        Me.DGV.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(17).HeaderText = "Nivel 6"
        Me.DGV.Columns(17).Width = 60
        Me.DGV.Columns(17).Visible = True
        Me.DGV.Columns(17).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(18).HeaderText = "Nivel 7"
        Me.DGV.Columns(18).Width = 60
        Me.DGV.Columns(18).Visible = True
        Me.DGV.Columns(18).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(19).HeaderText = "Nivel 8"
        Me.DGV.Columns(19).Width = 60
        Me.DGV.Columns(19).Visible = True
        Me.DGV.Columns(19).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(20).HeaderText = "FECHA_ALTA"
        Me.DGV.Columns(20).Width = 0
        Me.DGV.Columns(20).Visible = False

        Me.DGV.Columns(21).HeaderText = "USUARIO_EJECUTA_ALTA"
        Me.DGV.Columns(21).Width = 0
        Me.DGV.Columns(21).Visible = False

        Me.DGV.Columns(22).HeaderText = "TERMINAL_ALTA"
        Me.DGV.Columns(22).Width = 0
        Me.DGV.Columns(22).Visible = False

        Me.DGV.Columns(23).HeaderText = "IP_ALTA"
        Me.DGV.Columns(23).Width = 0
        Me.DGV.Columns(23).Visible = False

        Me.DGV.Columns(24).HeaderText = "FECHA_BAJA"
        Me.DGV.Columns(24).Width = 0
        Me.DGV.Columns(24).Visible = False

        Me.DGV.Columns(25).HeaderText = "USUARIO_EJECUTA_BAJA"
        Me.DGV.Columns(25).Width = 0
        Me.DGV.Columns(25).Visible = False

        Me.DGV.Columns(26).HeaderText = "TERMINAL_BAJA"
        Me.DGV.Columns(26).Width = 0
        Me.DGV.Columns(26).Visible = False

        Me.DGV.Columns(27).HeaderText = "IP_BAJA"
        Me.DGV.Columns(27).Width = 0
        Me.DGV.Columns(27).Visible = False

        Me.DGV.Columns(28).HeaderText = "ID_ESTADO_R"
        Me.DGV.Columns(28).Width = 0
        Me.DGV.Columns(28).Visible = False

    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        Dim I As Integer
        For I = 0 To Me.DGV.RowCount - 1
            If Me.DGV.Item(0, I).Value = uID_USUARIO Then
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
                    uID_USUARIO = Me.DGV.Item(0, I).Value
                    uID_TIPO_USUARIO = Me.DGV.Item(1, I).Value
                    uD_TIPO_USUARIO = Me.DGV.Item(2, I).Value
                    uUSUARIO = Me.DGV.Item(3, I).Value
                    uCUENTA = Me.DGV.Item(4, I).Value
                    uCLAVE = Me.DGV.Item(5, I).Value
                    uID_SESION_USUARIO = Me.DGV.Item(6, I).Value
                    uD_SESION_USUARIO = Me.DGV.Item(7, I).Value
                    uID_ESTADO_USUARIO = Me.DGV.Item(8, I).Value
                    uD_ESTADO_USUARIO = Me.DGV.Item(9, I).Value
                    If Not IsDBNull(Me.DGV.Item(10, I).Value) Then : uCODIGO = Me.DGV.Item(10, I).Value : Else : uCODIGO = "." : End If
                    If Not IsDBNull(Me.DGV.Item(11, I).Value) Then : uCARGO = Me.DGV.Item(11, I).Value : Else : uCARGO = "." : End If
                    If Not IsDBNull(Me.DGV.Item(12, I).Value) Then : uNIVEL1 = Me.DGV.Item(12, I).Value : Else : uNIVEL1 = "." : End If
                    If Not IsDBNull(Me.DGV.Item(13, I).Value) Then : uNIVEL2 = Me.DGV.Item(13, I).Value : Else : uNIVEL2 = "." : End If
                    If Not IsDBNull(Me.DGV.Item(14, I).Value) Then : uNIVEL3 = Me.DGV.Item(14, I).Value : Else : uNIVEL3 = "." : End If
                    If Not IsDBNull(Me.DGV.Item(15, I).Value) Then : uNIVEL4 = Me.DGV.Item(15, I).Value : Else : uNIVEL4 = "." : End If
                    If Not IsDBNull(Me.DGV.Item(16, I).Value) Then : uNIVEL5 = Me.DGV.Item(16, I).Value : Else : uNIVEL5 = "." : End If
                    If Not IsDBNull(Me.DGV.Item(17, I).Value) Then : uNIVEL6 = Me.DGV.Item(17, I).Value : Else : uNIVEL6 = "." : End If
                    If Not IsDBNull(Me.DGV.Item(18, I).Value) Then : uNIVEL7 = Me.DGV.Item(18, I).Value : Else : uNIVEL7 = "." : End If
                    If Not IsDBNull(Me.DGV.Item(19, I).Value) Then : uNIVEL8 = Me.DGV.Item(19, I).Value : Else : uNIVEL8 = "." : End If
                    If Not IsDBNull(Me.DGV.Item(20, I).Value) Then : uFECHA_ALTA = Me.DGV.Item(20, I).Value : Else : uFECHA_ALTA = "." : End If
                    If Not IsDBNull(Me.DGV.Item(21, I).Value) Then : uUSUARIO_EJECUTA_ALTA = Me.DGV.Item(21, I).Value : Else : uUSUARIO_EJECUTA_ALTA = "." : End If
                    If Not IsDBNull(Me.DGV.Item(22, I).Value) Then : uTERMINAL_ALTA = Me.DGV.Item(22, I).Value : Else : uTERMINAL_ALTA = "." : End If
                    If Not IsDBNull(Me.DGV.Item(23, I).Value) Then : uIP_ALTA = Me.DGV.Item(23, I).Value : Else : uIP_ALTA = "." : End If
                    If Not IsDBNull(Me.DGV.Item(24, I).Value) Then : uFECHA_BAJA = Me.DGV.Item(24, I).Value : Else : uFECHA_BAJA = "." : End If
                    If Not IsDBNull(Me.DGV.Item(25, I).Value) Then : uUSUARIO_EJECUTA_BAJA = Me.DGV.Item(25, I).Value : Else : uUSUARIO_EJECUTA_BAJA = "." : End If
                    If Not IsDBNull(Me.DGV.Item(26, I).Value) Then : uTERMINAL_BAJA = Me.DGV.Item(26, I).Value : Else : uTERMINAL_BAJA = "." : End If
                    If Not IsDBNull(Me.DGV.Item(27, I).Value) Then : uIP_BAJA = Me.DGV.Item(27, I).Value : Else : uIP_BAJA = "." : End If
                    If Not IsDBNull(Me.DGV.Item(28, I).Value) Then : uID_ESTADO_R = Me.DGV.Item(28, I).Value : Else : uID_ESTADO_R = "." : End If
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If uID_USUARIO = 0 Then
            MsgBox("Debe seleccionar el registro que desea Pasivar", vbInformation, "Mensaje del Sistema")
            Me.DGV.Focus()
            Exit Sub
        End If
        If uID_USUARIO = 1 Then
            MsgBox("El usuario Administrador no puede ser Apasivado", vbInformation, "Mensaje del Sistema")
            Me.DGV.Focus()
            Exit Sub
        End If
        If uID_ESTADO_R = 2 Then
            MsgBox("El registro se encuentra Pasivado y no se puede Editar", vbInformation, "Mensaje del Sistema")
            Me.DGV.Focus()
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Esta seguro de Pasivar el registro seleccionado?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            Call APASIVAR_MAESTRO_DE_NODOS_USUARIOS()
            Call APASIVAR_REGISTRO()
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            uID_USUARIO = 0
        Else
            Exit Sub
        End If
    End Sub
    Private Sub APASIVAR_MAESTRO_DE_NODOS_USUARIOS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE NODOS USUARIOS] SET ACCESO = 0 WHERE ID_USUARIO = " & CInt(uID_USUARIO) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub APASIVAR_REGISTRO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE USUARIOS] SET ID_SESION_USUARIO = 2, ID_ESTADO_USUARIO = 2, ID_ESTADO_R = 2 WHERE ID_USUARIO= " & CInt(uID_USUARIO) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub MovAtlaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MovAtlaToolStripMenuItem.Click
        If uID_USUARIO = 0 Then
            MsgBox("Debe seleccionar el Usuario al que le desea Establecer una Nueva Contraseña", vbInformation, "Mensaje del Sistema")
            Me.DGV.Focus()
            Exit Sub
        End If
        If uID_ESTADO_R = 2 Then
            MsgBox("El registro se encuentra Pasivado y no se puede Editar", vbInformation, "Mensaje del Sistema")
            Me.DGV.Focus()
            Exit Sub
        End If
        If uID_USUARIO <> 0 Then
            Frm_Establecer_Clave.ShowDialog()
            uID_USUARIO = 0
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Frm_Agregar_Usuarios.ShowDialog()
        If CERRAR = False Then
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
            uID_USUARIO = 0
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If uID_USUARIO = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Me.DGV.Focus()
            Exit Sub
        End If
        If uID_ESTADO_R = 2 Then
            MsgBox("El registro se encuentra Pasivado y no se puede Editar", vbInformation, "Mensaje del Sistema")
            Me.DGV.Focus()
            Exit Sub
        End If
        Frm_Editar_Usuarios.ShowDialog()
        If CERRAR = False Then
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
            uID_USUARIO = 0
        End If
    End Sub
    Private Sub DGV_DoubleClick(sender As Object, e As EventArgs) Handles DGV.DoubleClick
        If uID_USUARIO = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Me.DGV.Focus()
            Exit Sub
        End If
        If uID_ESTADO_R = 2 Then
            MsgBox("El registro se encuentra Pasivado y no se puede Editar", vbInformation, "Mensaje del Sistema")
            Me.DGV.Focus()
            Exit Sub
        End If
        Frm_Editar_Usuarios.ShowDialog()
        If CERRAR = False Then
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
            uID_USUARIO = 0
        End If
    End Sub
    Private Sub Frm_Usuarios_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        USUARIOS = False
    End Sub
    Private Sub Frm_Usuarios_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        USUARIOS = False
    End Sub
    Private Sub Frm_Usuarios_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        USUARIOS = False
    End Sub
    Private Sub Frm_Usuarios_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        USUARIOS = False
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        PARAMETRO = 1
        SELECCION = "({MAESTRO_DE_USUARIOS.ID_TIPO_USUARIO} = " & Me.ComboBox2.SelectedValue & ")"
        SELECCION_PARAMETRO = "*"
        USUARIO_IMPRIME = "IMPRESO POR: " & isuCUENTA
        Frm_Visor_Reportes.ShowDialog()
    End Sub
    Private Sub DGV_RowPrePaint(sender As Object, e As DataGridViewRowPrePaintEventArgs) Handles DGV.RowPrePaint
        Call FORMATO_FILAS_DATAGRIDVIEW()
    End Sub
    Private Sub FORMATO_FILAS_DATAGRIDVIEW()
        Dim style As New DataGridViewCellStyle
        style.Font = New Font(DGV.Font, FontStyle.Bold)
        For Each Row As DataGridViewRow In DGV.Rows
            If Row.Cells(28).Value = "2" Then
                Row.DefaultCellStyle.BackColor = Color.DimGray
            End If
        Next
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If uID_USUARIO = 0 Then
            MsgBox("Debe seleccionar el registro al que desea asignar Accesos", vbInformation, "Mensaje del Sistema")
            Me.DGV.Focus()
            Exit Sub
        End If
        Frm_Accesos.ShowDialog()
        uID_USUARIO = 0
    End Sub
End Class