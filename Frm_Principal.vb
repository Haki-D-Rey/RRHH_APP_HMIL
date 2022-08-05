Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Principal
    Private Sub Frm_Principal_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Frm_Seguridad.ShowDialog()
        If CERRAR = False Then
            Call OCULTAR_MENUS()
            Me.Timer1.Start()
            Me.Barra.Visible = True
            Me.T01.Text = "Terminal: " & My.Computer.Name
            Me.T02.Text = "Usuario Activo: " & isuCUENTA
            Me.T03.Text = "Tipo de Usuario: " & isuD_TIPO_USUARIO
            Me.PanelPrincipalA.Visible = True
            Me.PanelPrincipalB.Visible = True
            Me.CheckBox1.Checked = True
            Call VERIFICAR_SI_HAY_PROYECCION()
            If HAY_PROYECCION = True Then
                If isuD_TIPO_USUARIO <> "OPERADOR EXTERNO" Then
                    Frm_Principal_Menu_Emergente.Show(Me)
                End If
            End If
            If Me.CheckBox1.Checked = True Then
                tmMOSTRAR.Enabled = True
                Exit Sub
            End If
            Me.tmMOSTRAR.Enabled = True
        Else
            End
        End If
    End Sub
    Dim HAY_PROYECCION As Boolean
    Private Sub VERIFICAR_SI_HAY_PROYECCION()
        Call Clases.BUSCAR_FECHA_Y_HORA()
        FECHA01 = Mid(FECHA_SERVIDOR, 1, 10)
        Dim FECHA1 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA01 + "', 105), 23)"
        CADENAsql = "SELECT * FROM [dbo].[VISTA HISTORICO DE MOVIMIENTO] WHERE (FEC_APLICA > = " & FECHA1 & ") AND (ID_T_M_EXP = 1) AND (APLICAR = 'FALSE' AND ESTADISTICO = 'FALSE') ORDER BY FEC_APLICA"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                HAY_PROYECCION = True
            Else
                HAY_PROYECCION = False
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub OCULTAR_MENUS()
        Panel01.Visible = True
        Panel02.Visible = True
        Panel03.Visible = False
        Panel04.Visible = False
        Panel05.Visible = True
        Panel06.Visible = True
    End Sub
    Private Sub ShowSubmenu(SubMenu As Panel)
        If SubMenu.Visible = False Then
            SubMenu.Visible = True
        Else
            SubMenu.Visible = False
        End If
    End Sub
    'PROCESOS --------------------------------------------------
    Private Sub btnPROCESOS_Click(sender As Object, e As EventArgs) Handles btnPROCESOS.Click
        ShowSubmenu(Panel01)
    End Sub
    'Estructura Organizacional
    Private Sub button01A_Click(sender As Object, e As EventArgs) Handles button01A.Click
        Call VERIFICAR_ACCESOS("003") : If HAY_ACCESO = False Then : Exit Sub : End If
        Me.Cursor = Cursors.WaitCursor
        If ESTRUCTURA = True Then
            tmOCULTAR.Enabled = True
            Me.Timer1.Stop()
            Frm_Estructura_Org.WindowState = FormWindowState.Normal
            Me.Timer1.Start()
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        Me.Cursor = Cursors.Default
        tmOCULTAR.Enabled = True
        Me.Timer1.Stop()
        Frm_Estructura_Org.Show(Me)
        Me.Timer1.Start()
    End Sub
    'Movimientos
    Private Sub button02A_Click(sender As Object, e As EventArgs) Handles button02A.Click
        Call VERIFICAR_ACCESOS("004") : If HAY_ACCESO = False Then : Exit Sub : End If
        Me.Cursor = Cursors.WaitCursor
        If MOVIMIENTOS = True Then
            tmOCULTAR.Enabled = True
            Me.Timer1.Stop()
            Frm_Movimientos_Exp.WindowState = FormWindowState.Normal
            Me.Timer1.Start()
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        Me.Cursor = Cursors.Default
        tmOCULTAR.Enabled = True
        Me.Timer1.Stop()
        Frm_Movimientos_Exp.Show(Me)
        Me.Timer1.Start()
    End Sub
    'Aspirantes
    Private Sub button03A_Click(sender As Object, e As EventArgs) Handles button03A.Click
        Call VERIFICAR_ACCESOS("031") : If HAY_ACCESO = False Then : Exit Sub : End If
        Me.Cursor = Cursors.WaitCursor
        If ASPIRANTES = True Then
            tmOCULTAR.Enabled = True
            Me.Timer1.Stop()
            Frm_Aspirantes.WindowState = FormWindowState.Normal
            Me.Timer1.Start()
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        Me.Cursor = Cursors.Default
        tmOCULTAR.Enabled = True
        Me.Timer1.Stop()
        Frm_Aspirantes.Show(Me)
        Me.Timer1.Start()
    End Sub
    'Expedientes
    Private Sub button04A_Click(sender As Object, e As EventArgs) Handles button04A.Click
        Call VERIFICAR_ACCESOS("032") : If HAY_ACCESO = False Then : Exit Sub : End If
        Me.Cursor = Cursors.WaitCursor
        If EXPEDIENTES = True Then
            tmOCULTAR.Enabled = True
            Me.Timer1.Stop()
            Frm_Expedientes_Activos.WindowState = FormWindowState.Normal
            Me.Timer1.Start()
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        Me.Cursor = Cursors.Default
        tmOCULTAR.Enabled = True
        Me.Timer1.Stop()
        Frm_Expedientes_Activos.Show(Me)
        Me.Timer1.Start()
    End Sub
    'Medicos Acreditados
    Private Sub button05A_Click(sender As Object, e As EventArgs) Handles button05A.Click
        Call VERIFICAR_ACCESOS("033") : If HAY_ACCESO = False Then : Exit Sub : End If
        Me.Cursor = Cursors.WaitCursor
        If ACREDITADOS = True Then
            tmOCULTAR.Enabled = True
            Me.Timer1.Stop()
            Frm_Medicos_Acreditados.WindowState = FormWindowState.Normal
            Me.Timer1.Start()
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        Me.Cursor = Cursors.Default
        tmOCULTAR.Enabled = True
        Me.Timer1.Stop()
        Frm_Medicos_Acreditados.Show(Me)
        Me.Timer1.Start()
    End Sub
    'Expedientes Inactivos
    Private Sub button07A_Click(sender As Object, e As EventArgs) Handles button07A.Click
        Call VERIFICAR_ACCESOS("034") : If HAY_ACCESO = False Then : Exit Sub : End If
        Me.Cursor = Cursors.WaitCursor
        If EXPEDIENTES_I = True Then
            tmOCULTAR.Enabled = True
            Me.Timer1.Stop()
            Frm_Expedientes_Inactivos.WindowState = FormWindowState.Normal
            Me.Timer1.Start()
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        Me.Cursor = Cursors.Default
        tmOCULTAR.Enabled = True
        Me.Timer1.Stop()
        Frm_Expedientes_Inactivos.Show(Me)
        Me.Timer1.Start()
    End Sub
    'Situaciones del Personal
    Private Sub button08A_Click(sender As Object, e As EventArgs) Handles button08A.Click
        Call VERIFICAR_ACCESOS("035") : If HAY_ACCESO = False Then : Exit Sub : End If
        Me.Cursor = Cursors.WaitCursor
        If SITUACIONES = True Then
            tmOCULTAR.Enabled = True
            Me.Timer1.Stop()
            Frm_Situacion_Personal.WindowState = FormWindowState.Normal
            Me.Timer1.Start()
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        Me.Cursor = Cursors.Default
        tmOCULTAR.Enabled = True
        Me.Timer1.Stop()
        Frm_Situacion_Personal.Show(Me)
        Me.Timer1.Start()
    End Sub
    'Accidentes Laborales
    Private Sub button10A_Click(sender As Object, e As EventArgs) Handles button10A.Click
        Call VERIFICAR_ACCESOS("036") : If HAY_ACCESO = False Then : Exit Sub : End If
        Me.Cursor = Cursors.WaitCursor
        If HIGIENE_1 = True Then
            tmOCULTAR.Enabled = True
            Me.Timer1.Stop()
            Frm_Accidentes_Laborales.WindowState = FormWindowState.Normal
            Me.Timer1.Start()
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        Me.Cursor = Cursors.Default
        tmOCULTAR.Enabled = True
        Me.Timer1.Stop()
        Frm_Accidentes_Laborales.Show(Me)
        Me.Timer1.Start()
    End Sub
    'Dosimetria
    Private Sub button11A_Click(sender As Object, e As EventArgs) Handles button11A.Click
        Call VERIFICAR_ACCESOS("037") : If HAY_ACCESO = False Then : Exit Sub : End If
        Me.Cursor = Cursors.WaitCursor
        If HIGIENE_2 = True Then
            tmOCULTAR.Enabled = True
            Me.Timer1.Stop()
            Frm_Dosimetria.WindowState = FormWindowState.Normal
            Me.Timer1.Start()
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        Me.Cursor = Cursors.Default
        tmOCULTAR.Enabled = True
        Me.Timer1.Stop()
        Frm_Dosimetria.Show(Me)
        Me.Timer1.Start()
    End Sub
    'Examenes Medicos Ocupacionales
    Private Sub button12A_Click(sender As Object, e As EventArgs) Handles button12A.Click
        Call VERIFICAR_ACCESOS("038") : If HAY_ACCESO = False Then : Exit Sub : End If
        Me.Cursor = Cursors.WaitCursor
        If HIGIENE_3 = True Then
            tmOCULTAR.Enabled = True
            Me.Timer1.Stop()
            Frm_Ex_Med_Ocup.WindowState = FormWindowState.Normal
            Me.Timer1.Start()
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        Me.Cursor = Cursors.Default
        tmOCULTAR.Enabled = True
        Me.Timer1.Stop()
        Frm_Ex_Med_Ocup.Show(Me)
        Me.Timer1.Start()
    End Sub
    'Personal Tercerizado
    Private Sub button14A_Click(sender As Object, e As EventArgs) Handles button14A.Click
        Call VERIFICAR_ACCESOS("039") : If HAY_ACCESO = False Then : Exit Sub : End If
        Me.Cursor = Cursors.WaitCursor
        If TERCERIZADOS = True Then
            tmOCULTAR.Enabled = True
            Me.Timer1.Stop()
            Frm_Personal_Tercerizado.WindowState = FormWindowState.Normal
            Me.Timer1.Start()
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        Me.Cursor = Cursors.Default
        tmOCULTAR.Enabled = True
        Me.Timer1.Stop()
        Frm_Personal_Tercerizado.Show(Me)
        Me.Timer1.Start()
    End Sub
    'Becados
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call VERIFICAR_ACCESOS("040") : If HAY_ACCESO = False Then : Exit Sub : End If
        Me.Cursor = Cursors.WaitCursor
        If BECADOS = True Then
            tmOCULTAR.Enabled = True
            Me.Timer1.Stop()
            Frm_Becados.WindowState = FormWindowState.Normal
            Me.Timer1.Start()
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        Me.Cursor = Cursors.Default
        tmOCULTAR.Enabled = True
        Me.Timer1.Stop()
        Frm_Becados.Show(Me)
        Me.Timer1.Start()
    End Sub
    'CONSULTAS --------------------------------------------------
    'Selección de Consultas y Reportes
    Private Sub button01B_Click(sender As Object, e As EventArgs) Handles button01B.Click
        Call VERIFICAR_ACCESOS("041") : If HAY_ACCESO = False Then : Exit Sub : End If
        Me.Cursor = Cursors.WaitCursor
        If INFORMES = True Then
            tmOCULTAR.Enabled = True
            Me.Timer1.Stop()
            Frm_Seleccion_Consultas_y_Reportes.WindowState = FormWindowState.Normal
            Me.Timer1.Start()
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        Me.Cursor = Cursors.Default
        tmOCULTAR.Enabled = True
        Me.Timer1.Stop()
        Frm_Seleccion_Consultas_y_Reportes.Show(Me)
        Me.Timer1.Start()
    End Sub
    'Consulta General
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call VERIFICAR_ACCESOS("042") : If HAY_ACCESO = False Then : Exit Sub : End If
        Me.Cursor = Cursors.WaitCursor
        If CONSULTA_GENERAL = True Then
            tmOCULTAR.Enabled = True
            Me.Timer1.Stop()
            Frm_Consulta_General.WindowState = FormWindowState.Normal
            Me.Timer1.Start()
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        Me.Cursor = Cursors.Default
        tmOCULTAR.Enabled = True
        Me.Timer1.Stop()
        Frm_Consulta_General.Show(Me)
        Me.Timer1.Start()
    End Sub
    'Consulta Colaboradores con Saldos Vac.
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call VERIFICAR_ACCESOS("043") : If HAY_ACCESO = False Then : Exit Sub : End If
        Me.Cursor = Cursors.WaitCursor
        If CONSULTA_SAL_VAC = True Then
            tmOCULTAR.Enabled = True
            Me.Timer1.Stop()
            Frm_Consulta_Saldos.WindowState = FormWindowState.Normal
            Me.Timer1.Start()
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        Me.Cursor = Cursors.Default
        tmOCULTAR.Enabled = True
        Me.Timer1.Stop()
        Frm_Consulta_Saldos.Show(Me)
        Me.Timer1.Start()
    End Sub
    'Matriz BAC
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Call VERIFICAR_ACCESOS("044") : If HAY_ACCESO = False Then : Exit Sub : End If
        Me.Cursor = Cursors.WaitCursor
        If CONSULTA_MATRIZ_BAC = True Then
            tmOCULTAR.Enabled = True
            Me.Timer1.Stop()
            Frm_Consulta_Matriz_BAC.WindowState = FormWindowState.Normal
            Me.Timer1.Start()
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        Me.Cursor = Cursors.Default
        tmOCULTAR.Enabled = True
        Me.Timer1.Stop()
        Frm_Consulta_Matriz_BAC.Show(Me)
        Me.Timer1.Start()
    End Sub
    'Personal Ausente en 4 (cuatro) Dias continuos
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Call VERIFICAR_ACCESOS("045") : If HAY_ACCESO = False Then : Exit Sub : End If
        Me.Cursor = Cursors.WaitCursor
        If CONSULTA_PERSONAL_AUSENTE = True Then
            tmOCULTAR.Enabled = True
            Me.Timer1.Stop()
            Frm_Personal_Ausente_Dias_Continuos.WindowState = FormWindowState.Normal
            Me.Timer1.Start()
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        Me.Cursor = Cursors.Default
        tmOCULTAR.Enabled = True
        Me.Timer1.Stop()
        Frm_Personal_Ausente_Dias_Continuos.Show(Me)
        Me.Timer1.Start()
    End Sub
    'Cobertura Militar\ Afiliados
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Call VERIFICAR_ACCESOS("046") : If HAY_ACCESO = False Then : Exit Sub : End If
        Me.Cursor = Cursors.WaitCursor
        If COBERTURA_AFILIADOS = True Then
            tmOCULTAR.Enabled = True
            Me.Timer1.Stop()
            Frm_Verificar.WindowState = FormWindowState.Normal
            Me.Timer1.Start()
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        Me.Cursor = Cursors.Default
        tmOCULTAR.Enabled = True
        Me.Timer1.Stop()
        Frm_Verificar.Show(Me)
        Me.Timer1.Start()
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Call VERIFICAR_ACCESOS("309") : If HAY_ACCESO = False Then : Exit Sub : End If
        Me.Cursor = Cursors.WaitCursor
        If TIEMPO_EXTRAORDINARIO = True Then
            tmOCULTAR.Enabled = True
            Me.Timer1.Stop()
            Frm_Tiempo_Extraordinario.WindowState = FormWindowState.Normal
            Me.Timer1.Start()
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        Me.Cursor = Cursors.Default
        tmOCULTAR.Enabled = True
        Me.Timer1.Stop()
        Frm_Tiempo_Extraordinario.Show(Me)
        Me.Timer1.Start()
    End Sub
    Private Sub btnCONSULTAS_Click(sender As Object, e As EventArgs) Handles btnCONSULTAS.Click
        ShowSubmenu(Panel02)
    End Sub
    Private Sub btnCATALOGOS_Click(sender As Object, e As EventArgs) Handles btnCATALOGOS.Click
        ShowSubmenu(Panel03)
    End Sub
    'CONTROL DE CATALOGOS --------------------------------------------------
    'Catalogos
    Private Sub button01C_Click(sender As Object, e As EventArgs) Handles button01C.Click
        Call VERIFICAR_ACCESOS("020") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Catalogos_Seleccion.TopMost = False
        Frm_Catalogos_Seleccion.TopLevel = False
        Frm_Catalogos_Seleccion.Parent = Me.PanelPrincipalB
        Frm_Catalogos_Seleccion.Dock = DockStyle.Fill
        'Frm_Catalogos_Seleccion.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        Frm_Catalogos_Seleccion.BackColor = Color.White
        Frm_Catalogos_Seleccion.Show()
    End Sub
    Private Sub btnUTILIDADES_Click(sender As Object, e As EventArgs) Handles btnUTILIDADES.Click
        ShowSubmenu(Panel04)
    End Sub
    'Usuarios\ Accesos
    Private Sub button01D_Click(sender As Object, e As EventArgs) Handles button01D.Click
        Call VERIFICAR_ACCESOS("022") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Usuarios.ShowDialog()
    End Sub
    'Configuraciones
    Private Sub button02D_Click(sender As Object, e As EventArgs) Handles button02D.Click
        Call VERIFICAR_ACCESOS("023") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Configuraciones.ShowDialog()
    End Sub
    'Carga de Marcas Biometricas
    Private Sub button03D_Click(sender As Object, e As EventArgs) Handles button03D.Click
        Call VERIFICAR_ACCESOS("047") : If HAY_ACCESO = False Then : Exit Sub : End If
        Call CARGAR_MARCAS()
        MsgBox("Carga de Marcas Biometricas Finalizada", vbInformation, "Mensaje del Sistema")
    End Sub
    'Actualizacion de Documentacion Digital
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Call VERIFICAR_ACCESOS("048") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Act_Documentos_Digitales.ShowDialog()
    End Sub
    'OTROS --------------------------------------------------
    'Estadisticas
    Private Sub button01E_Click(sender As Object, e As EventArgs) Handles button01E.Click
        Call VERIFICAR_ACCESOS("049") : If HAY_ACCESO = False Then : Exit Sub : End If
        Me.Cursor = Cursors.WaitCursor
        If ESTADISTICAS = True Then
            Me.Timer1.Stop()
            Frm_Estadisticas.WindowState = FormWindowState.Normal
            Me.Timer1.Start()
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        Me.Cursor = Cursors.Default
        Me.Timer1.Stop()
        Frm_Estadisticas.Show(Me)
        Me.Timer1.Start()
    End Sub
    Private Sub btnOTROS_Click(sender As Object, e As EventArgs) Handles btnOTROS.Click
        ShowSubmenu(Panel05)
    End Sub
    Private Sub btnSISTEMAS_Click(sender As Object, e As EventArgs) Handles btnSISTEMAS.Click
        ShowSubmenu(Panel06)
    End Sub
    'Cambiar clave de Usuario
    Private Sub button01F_Click(sender As Object, e As EventArgs) Handles button01F.Click
        Call VERIFICAR_ACCESOS("028") : If HAY_ACCESO = False Then : Exit Sub : End If
        Me.Timer1.Stop()
        Frm_Cambiar_Clave.ShowDialog()
        If CERRAR = False Then
            Call Clases.BUSCAR_FECHA_Y_HORA()
            vIP = Clases.GetIP()
            MsgBox("El sistema Finalizara la sesion Actual debido al cambio de Contraseña de Usuario", vbInformation, "Mensaje del Sistema")
            Call ACTUALIZAR_ACCESO_DE_USUARIO_SALIR()
            Me.Barra.Visible = False
            Me.Timer1.Stop()
            Me.PanelPrincipalA.Visible = False
            Me.PanelPrincipalB.Visible = False
            Frm_Seguridad.ShowDialog()
            If CERRAR = False Then
                Call Clases.BUSCAR_FECHA_Y_HORA()
                vIP = Clases.GetIP()
                Call OCULTAR_MENUS()
                Me.Timer1.Start()
                Me.Barra.Visible = True
                Me.T01.Text = "Terminal: " & My.Computer.Name
                Me.T02.Text = "Usuario Activo: " & isuCUENTA
                Me.T03.Text = "Tipo de Usuario: " & isuD_TIPO_USUARIO
                Me.PanelPrincipalA.Visible = True
                Me.PanelPrincipalB.Visible = True
                Call VERIFICAR_SI_HAY_PROYECCION()
                If HAY_PROYECCION = True Then
                    If isuD_TIPO_USUARIO <> "OPERADOR EXTERNO" Then
                        Frm_Principal_Menu_Emergente.Show(Me)
                    End If
                End If
            Else
                End
            End If
        Else
            Me.Timer1.Start()
        End If
    End Sub
    'Cerrar Sesion
    Private Sub button02F_Click(sender As Object, e As EventArgs) Handles button02F.Click
        Call VERIFICAR_ACCESOS("029") : If HAY_ACCESO = False Then : Exit Sub : End If
        Dim cadena As String = MessageBox.Show("¿Desea Cerrar la Sesion Actual del Sistema?", "Mensaje del Sistema", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2)
        If cadena = vbYes Then
            Call Clases.BUSCAR_FECHA_Y_HORA()
            vIP = Clases.GetIP()
            Call ACTUALIZAR_ACCESO_DE_USUARIO_SALIR()
            Me.Barra.Visible = False
            Me.Timer1.Stop()
            Me.PanelPrincipalA.Visible = False
            Me.PanelPrincipalB.Visible = False
            Frm_Seguridad.ShowDialog()
            If CERRAR = False Then
                Call Clases.BUSCAR_FECHA_Y_HORA()
                vIP = Clases.GetIP()
                Call OCULTAR_MENUS()
                Me.Timer1.Start()
                Me.Barra.Visible = True
                Me.T01.Text = "Terminal: " & My.Computer.Name
                Me.T02.Text = "Usuario Activo: " & isuCUENTA
                Me.T03.Text = "Tipo de Usuario: " & isuD_TIPO_USUARIO
                Me.PanelPrincipalA.Visible = True
                Call VERIFICAR_SI_HAY_PROYECCION()
                If HAY_PROYECCION = True Then
                    If isuD_TIPO_USUARIO <> "OPERADOR EXTERNO" Then
                        Frm_Principal_Menu_Emergente.Show(Me)
                    End If
                End If
            Else
                End
            End If
        End If
    End Sub
    'Salir
    Private Sub button03F_Click(sender As Object, e As EventArgs) Handles button03F.Click
        Call VERIFICAR_ACCESOS("030") : If HAY_ACCESO = False Then : Exit Sub : End If
        Dim cadena As String = MessageBox.Show("¿Desea salir del Sistema?", "Mensaje del Sistema", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2)
        If cadena = vbYes Then
            Call ACTUALIZAR_ACCESO_DE_USUARIO_SALIR()
            End
        End If
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Call Clases.BUSCAR_FECHA_Y_HORA()
        Me.T05.Text = "Fecha Actual: " & Mid(FECHA_SERVIDOR, 1, 10)
        Me.T06.Text = "Hora Actual: " & HORA_SERVIDOR.Hour & ":" & HORA_SERVIDOR.Minute & ":" & HORA_SERVIDOR.Second
    End Sub
    Private Sub ACTUALIZAR_ACCESO_DE_USUARIO_SALIR()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE USUARIOS] SET ID_SESION_USUARIO = '2' WHERE ID_USUARIO = " & CInt(isuID_USUARIO) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Frm_Principal_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Call ACTUALIZAR_ACCESO_DE_USUARIO_SALIR()
    End Sub
    '**********************************************************************
    '**********************************************************************
    '**********************************************************************
    '**********************************************************************
    'PROCESO DE CARGA DE MARCAS BIOMETRICAS
    Dim MDT As New DataTable
    Dim MDA As SqlDataAdapter
    Dim iMARC As Integer
    Private Sub CARGAR_MARCAS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Try
            CADENAsql = "Select [ID], [FECHA], Convert(VARCHAR(10), convert(Date, FECHA, 105),23) As FECHA_1, CODIGO, CHECKTYPE, TERMINAL " &
                        "from [dbo].[VISTA DE MARCAS EN BD MARCAS] A " &
                        "Where Not EXISTS " &
                        "(SELECT HORA_MARCA, ID_EMPLEADO " &
                        "From [dbo].[MAESTRO DE MARCAS] B " &
                        "Where A.[FECHA] = B.[HORA_MARCA] " &
                        "And A.[CODIGO] = B.[ID_EMPLEADO])"
            Conexion.ABD_RRHH()
            MDA = New SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
            MDT.Clear()
            MDA.Fill(MDT)
            If MDT.Rows.Count > 0 Then
                For iMARC = 0 To MDT.Rows.Count - 1
                    Application.DoEvents()
                    Me.Cursor = Cursors.WaitCursor
                    Call GRABAR_REGISTRO()
                    Me.Cursor = Cursors.Default
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
    Dim FECHA1 As DateTime
    Dim FECHA2 As DateTime
    Dim FECHA3 As DateTime
    Dim fHORA_MARCA As String
    Dim fFECHA_MARCA As String
    Dim fFECHA_CARGA As String
    Dim TIPO As String
    Private Sub GRABAR_REGISTRO()
        'I: ENTRADA, O: SALIDA 
        If MDT.Rows(iMARC).Item(5).ToString = "1" Then          'LOBBY (B3-PB)              'CONFIG CORRECTA AL 07/04/2021
            If MDT.Rows(iMARC).Item(4).ToString() = "O" Then
                TIPO = "ENTRADA"
            End If
            If MDT.Rows(iMARC).Item(4).ToString() = "I" Then
                TIPO = "SALIDA"
            End If
        End If
        'O: ENTRADA, 1: SALIDA
        If MDT.Rows(iMARC).Item(5).ToString = "3" Then          'EDIFICIO 8 (B8-P1) 1       'CONFIG CORRECTA AL 07/04/2021
            If MDT.Rows(iMARC).Item(4).ToString() = "O" Then
                TIPO = "ENTRADA"
            End If
            If MDT.Rows(iMARC).Item(4).ToString() = "1" Then
                TIPO = "SALIDA"
            End If
        End If
        'O: ENTRADA, 1: SALIDA
        If MDT.Rows(iMARC).Item(5).ToString = "4" Then          'FARMACIA (B1-PB)           'CONFIG CORRECTA AL 07/04/2021
            If MDT.Rows(iMARC).Item(4).ToString() = "O" Then
                TIPO = "ENTRADA"
            End If
            If MDT.Rows(iMARC).Item(4).ToString() = "I" Then
                TIPO = "SALIDA"
            End If
        End If
        'O: ENTRADA, I: SALIDA
        If MDT.Rows(iMARC).Item(5).ToString = "6" Then          'POL. ALTAMIRA (PB)         'CONFIG CORRECTA AL 07/04/2021
            If MDT.Rows(iMARC).Item(4).ToString() = "O" Then
                TIPO = "ENTRADA"
            End If
            If MDT.Rows(iMARC).Item(4).ToString() = "I" Then
                TIPO = "SALIDA"
            End If
        End If
        'O: ENTRADA, i: SALIDA
        If MDT.Rows(iMARC).Item(5).ToString = "7" Then          'POL. TIPITAPA (PB)         'CONFIG CORRECTA AL 01/06/2021
            If MDT.Rows(iMARC).Item(4).ToString() = "O" Then
                TIPO = "ENTRADA"
            End If
            If MDT.Rows(iMARC).Item(4).ToString() = "i" Then
                TIPO = "SALIDA"
            End If
        End If
        Me.DateTimePicker1.Text = MDT.Rows(iMARC).Item(1).ToString()
        Me.DateTimePicker2.Text = MDT.Rows(iMARC).Item(1).ToString()
        Me.DateTimePicker3.Text = Now
        FECHA1 = Me.DateTimePicker1.Text
        FECHA2 = Me.DateTimePicker2.Text
        FECHA3 = Me.DateTimePicker3.Text
        Dim fHORA_MARCA As String = FECHA1.Year & "/" & FECHA1.Month & "/" & FECHA1.Day & " " & FECHA1.Hour & ":" & FECHA1.Minute & ":" & FECHA1.Second
        Dim fFECHA_MARCA As String = FECHA2.Year & "/" & FECHA2.Month & "/" & FECHA2.Day
        Dim fFECHA_CARGA As String = FECHA3.Year & "/" & FECHA3.Month & "/" & FECHA3.Day & " " & FECHA3.Hour & ":" & FECHA3.Minute
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim SQL As String
        Try
            SQL = "INSERT INTO [MAESTRO DE MARCAS] (IDEMPRESA, IDRELOJ, ID_EMPLEADO, HORA_MARCA, FECHA_MARCA, FECHA_CARGA, ARCHIVO, NUMERO, PROGRAMADO, OBSERVACION,TIPO_MARCA)" &
                                           "values ('1', '" & MDT.Rows(iMARC).Item(5).ToString & "', '" & MDT.Rows(iMARC).Item(3).ToString & "', '" & fHORA_MARCA & "', '" & fFECHA_MARCA & "', '" & fFECHA_CARGA & "', '.', '.', '0', '.', '" & TIPO & "')"
            Dim cmd As New SqlCommand(SQL, CxRRHH)
            cmd.ExecuteNonQuery()
            cmd = Nothing
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    '**********************************************************************
    '**********************************************************************
    '**********************************************************************
    '**********************************************************************
    Private Sub PanelPrincipalA_MouseHover(sender As Object, e As EventArgs) Handles PanelPrincipalA.MouseHover
        If Me.CheckBox1.Checked = True Then
            tmMOSTRAR.Enabled = True
            Exit Sub
        End If
        If PanelPrincipalA.Width = 339 Then
            tmOCULTAR.Enabled = True
        ElseIf PanelPrincipalA.Width = 40 Then
            tmMOSTRAR.Enabled = True
        End If
    End Sub
    Private Sub tmOCULTAR_Tick(sender As Object, e As EventArgs) Handles tmOCULTAR.Tick
        If Me.CheckBox1.Checked = True Then
            tmMOSTRAR.Enabled = True
            Exit Sub
        End If
        If PanelPrincipalA.Width <= 40 Then
            Me.tmOCULTAR.Enabled = False
        Else
            Me.PanelPrincipalA.Width = PanelPrincipalA.Width - 40
        End If
    End Sub
    Private Sub tmMOSTRAR_Tick(sender As Object, e As EventArgs) Handles tmMOSTRAR.Tick
        If Me.CheckBox1.Checked = True Then
            tmMOSTRAR.Enabled = True
            Exit Sub
        End If
        If PanelPrincipalA.Width >= 339 Then
            Me.tmMOSTRAR.Enabled = False
        Else
            Me.PanelPrincipalA.Width = PanelPrincipalA.Width + 40
        End If
    End Sub
    Private Sub PictureBox1_MouseHover(sender As Object, e As EventArgs) Handles PictureBox1.MouseHover
        If Me.CheckBox1.Checked = True Then
            tmMOSTRAR.Enabled = True
            Exit Sub
        End If
        tmOCULTAR.Enabled = True
    End Sub
    Private Sub Panel1_MouseHover(sender As Object, e As EventArgs) Handles Panel1.MouseHover
        If Me.CheckBox1.Checked = True Then
            tmMOSTRAR.Enabled = True
            Exit Sub
        End If
        Me.tmMOSTRAR.Enabled = True
    End Sub
    Private Sub GroupBox1_MouseHover(sender As Object, e As EventArgs) Handles GroupBox1.MouseHover
        If Me.CheckBox1.Checked = True Then
            tmMOSTRAR.Enabled = True
            Exit Sub
        End If
        Me.tmMOSTRAR.Enabled = True
    End Sub
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        If Me.CheckBox1.Checked = True Then
            tmMOSTRAR.Enabled = True
            Exit Sub
        End If
        tmOCULTAR.Enabled = True
    End Sub
    Private Sub Frm_Principal_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.F12) Then
            Frm_Users_Activos.TopMost = False
            Frm_Users_Activos.TopLevel = False
            Frm_Users_Activos.Parent = Me.PanelPrincipalB
            Frm_Users_Activos.Dock = DockStyle.Fill
            'Frm_Users_Activos.FormBorderStyle = Windows.Forms.FormBorderStyle.None
            Frm_Users_Activos.BackColor = Color.White
            Frm_Users_Activos.Show()
        End If
    End Sub
End Class