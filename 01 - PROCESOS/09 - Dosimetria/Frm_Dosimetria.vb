Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Public Class Frm_Dosimetria
    Private Sub Frm_Dosimetria_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            HIGIENE_2 = False
            Me.TextBox1.Text = ""
            Me.TextBox2.Text = ""
            Me.TextBox3.Text = ""
            Me.TextBox4.Text = ""
            Me.TextBox5.Text = ""
            Me.TextBox6.Text = ""
            Me.TextBox8.Text = ""
            Me.TextBox36.Text = "0"
            Me.TextBox2.Focus()
            Frm_Principal.Timer1.Start()
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Dosimetria_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Frm_Principal.Timer1.Stop()
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Me.TextBox6.Text = ""
        Me.TextBox36.Text = "0"
        Me.TextBox8.Text = ""
        CADENAsql = "Select * from [dbo].[VISTA MAESTRO DE HIGIENE DOSIMETRIA] WHERE ID_M_P = " & Me.TextBox36.Text & " ORDER BY FECHA_LECTURA"
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
        Me.TextBox2.Focus()
    End Sub
    Private Sub Button7_Click(sender As Object, e As System.EventArgs) Handles Button7.Click
        HIGIENE_2 = False
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Me.TextBox6.Text = ""
        Me.TextBox36.Text = "0"
        Me.TextBox8.Text = ""
        Me.TextBox2.Focus()
        Frm_Principal.Timer1.Start()
        Me.Close()
    End Sub
    Private Sub TextBox2_DoubleClick(sender As Object, e As EventArgs) Handles TextBox2.DoubleClick
        Call VERIFICAR_ACCESOS("211") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Accidentes_Laborales_Buscar.ShowDialog()
        CE = h1CODIGO
        If CE <> "" Then
            Me.TextBox2.Text = CE
            vID_M_P = h1ID_M_P
            Me.TextBox36.Text = vID_M_P
            CADENAsql = "SELECT ID_M_P, CODIGO, APELLIDOS, NOMBRES, N_CEDULA, FEC_ING_PAME, ID_ESTADO_P, ID_N_PROFESIONAL, D_T_SEXO FROM [dbo].[VISTA MAESTRO DE PERSONAS] WHERE CODIGO = '" & Me.TextBox2.Text & "'"
            Call BUSCAR_DATO_1()
            If DATO_1 = True Then
                CADENAsql = "Select * from [dbo].[VISTA MAESTRO DE HIGIENE DOSIMETRIA] WHERE ID_M_P = " & Me.TextBox36.Text & " ORDER BY FECHA_LECTURA"
                Call CARGAR_DATAGRIDVIEW()
                Call ARMAR_DATAGRIDVIEW()
                Call BUSCAR_CARGO()
            Else
                MsgBox("El código que se digitó no es válido", vbInformation, "Mensaje del Sistema")
                Me.TextBox1.SelectAll()
                Me.TextBox1.Focus()
                Exit Sub
            End If
            Me.TextBox2.SelectAll()
            Me.TextBox2.Focus()

        Else
            Me.TextBox2.Text = "0000000000"
            Me.TextBox36.Text = "0"
            Me.TextBox2.SelectAll()
            Me.TextBox2.Focus()
        End If
    End Sub
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call VERIFICAR_ACCESOS("211") : If HAY_ACCESO = False Then : Exit Sub : End If
            cEMPLEADO = CInt(Val(TextBox2.Text))
            TextBox2.Text = Format(CInt(cEMPLEADO), "0000000000")
            If Len(Me.TextBox2.Text) = 0 Or (Me.TextBox2.Text = "0000000000") Then
                MsgBox("El código que se digitó no es válido", vbInformation, "Mensaje del Sistema")
                Me.TextBox2.Focus()
                Exit Sub
            Else
                CADENAsql = "SELECT ID_M_P, CODIGO, APELLIDOS, NOMBRES, N_CEDULA, FEC_ING_PAME, ID_ESTADO_P, ID_N_PROFESIONAL, D_T_SEXO FROM [dbo].[VISTA MAESTRO DE PERSONAS] WHERE CODIGO = '" & Me.TextBox2.Text & "'"
                Call BUSCAR_DATO_1()
                If DATO_1 = True Then
                    CADENAsql = "Select * from [dbo].[VISTA MAESTRO DE HIGIENE DOSIMETRIA] WHERE ID_M_P = " & Me.TextBox36.Text & " ORDER BY FECHA_LECTURA"
                    Call CARGAR_DATAGRIDVIEW()
                    Call ARMAR_DATAGRIDVIEW()
                    Call BUSCAR_CARGO()
                    Me.TextBox2.SelectAll()
                    Me.TextBox2.Focus()
                Else
                    MsgBox("El código que se digitó no es válido", vbInformation, "Mensaje del Sistema")
                    Me.TextBox2.SelectAll()
                    Me.TextBox2.Focus()
                    Exit Sub
                End If
            End If
        End If
    End Sub
    Dim DE, SE As String
    Private Sub BUSCAR_CARGO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT ID_M_P, D_N1, D_N2, D_N3, D_N4, D_N5, D_N6, D_N7, D_N8, D_CARGO_ES, ID_SITUACION FROM [dbo].[VISTA MAESTRO DE CARGOS SIT] WHERE CODIGO = '" & Me.TextBox2.Text & "' AND ID_SITUACION = 1 ORDER BY ID_M_P", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                '---------------------------------------------------------------------------------- SERVICIO
                If MiDataTable.Rows(0).Item(8).ToString <> "" Then  'NIVEL 8
                    SE = MiDataTable.Rows(0).Item(8).ToString
                    GoTo SERVICIO
                Else
                    If MiDataTable.Rows(0).Item(7).ToString <> "" Then  'NIVEL 7
                        SE = MiDataTable.Rows(0).Item(7).ToString
                        GoTo SERVICIO
                    Else
                        If MiDataTable.Rows(0).Item(6).ToString <> "" Then   'NIVEL 6
                            SE = MiDataTable.Rows(0).Item(6).ToString
                            GoTo SERVICIO
                        Else
                            If MiDataTable.Rows(0).Item(5).ToString <> "" Then  'NIVEL 5
                                SE = MiDataTable.Rows(0).Item(5).ToString
                                GoTo SERVICIO
                            Else
                                If MiDataTable.Rows(0).Item(4).ToString <> "" Then  'NIVEL 4
                                    SE = MiDataTable.Rows(0).Item(4).ToString
                                    GoTo SERVICIO
                                Else
                                    If MiDataTable.Rows(0).Item(3).ToString <> "" Then  'NIVEL 3
                                        SE = MiDataTable.Rows(0).Item(3).ToString
                                        GoTo SERVICIO
                                    Else
                                        If MiDataTable.Rows(0).Item(2).ToString <> "" Then  'NIVEL 2
                                            SE = MiDataTable.Rows(0).Item(2).ToString
                                            GoTo SERVICIO
                                        Else
                                            If MiDataTable.Rows(0).Item(1).ToString <> "" Then  'NIVEL 1
                                                SE = MiDataTable.Rows(0).Item(1).ToString
                                                GoTo SERVICIO
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
SERVICIO:
                '---------------------------------------------------------------------------------- DEPARTAMENTO
                If MiDataTable.Rows(0).Item(8).ToString <> "" Then  'NIVEL 8
                    DE = MiDataTable.Rows(0).Item(8).ToString
                    If DE <> SE Then
                        GoTo DEPARTAMENTO
                    Else
                        GoTo LEE_1
                    End If
                Else
LEE_1:
                    If MiDataTable.Rows(0).Item(7).ToString <> "" Then  'NIVEL 7
                        DE = MiDataTable.Rows(0).Item(7).ToString
                        If DE <> SE Then
                            GoTo DEPARTAMENTO
                        Else
                            GoTo LEE_2
                        End If
                    Else
LEE_2:
                        If MiDataTable.Rows(0).Item(6).ToString <> "" Then  'NIVEL 6
                            DE = MiDataTable.Rows(0).Item(6).ToString
                            If DE <> SE Then
                                GoTo DEPARTAMENTO
                            Else
                                GoTo LEE_3
                            End If
                        Else
LEE_3:
                            If MiDataTable.Rows(0).Item(5).ToString <> "" Then  'NIVEL 5
                                DE = MiDataTable.Rows(0).Item(5).ToString
                                If DE <> SE Then
                                    GoTo DEPARTAMENTO
                                Else
                                    GoTo LEE_4
                                End If
                            Else
LEE_4:
                                If MiDataTable.Rows(0).Item(4).ToString <> "" Then  'NIVEL 4
                                    DE = MiDataTable.Rows(0).Item(4).ToString
                                    If DE <> SE Then
                                        GoTo DEPARTAMENTO
                                    Else
                                        GoTo LEE_5
                                    End If
                                Else
LEE_5:
                                    If MiDataTable.Rows(0).Item(3).ToString <> "" Then  'NIVEL 3
                                        DE = MiDataTable.Rows(0).Item(3).ToString
                                        If DE <> SE Then
                                            GoTo DEPARTAMENTO
                                        Else
                                            GoTo LEE_6
                                        End If
                                    Else
LEE_6:
                                        If MiDataTable.Rows(0).Item(2).ToString <> "" Then  'NIVEL 2
                                            DE = MiDataTable.Rows(0).Item(2).ToString
                                            If DE <> SE Then
                                                GoTo DEPARTAMENTO
                                            Else
                                                GoTo LEE_7
                                            End If
                                        Else
LEE_7:
                                            If MiDataTable.Rows(0).Item(1).ToString <> "" Then  'NIVEL 1
                                                DE = MiDataTable.Rows(0).Item(1).ToString
                                                If DE <> SE Then
                                                    GoTo DEPARTAMENTO
                                                Else
                                                    GoTo LEE_8
                                                End If
                                            End If
LEE_8:
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
DEPARTAMENTO:
                Me.TextBox4.Text = MiDataTable.Rows(0).Item(9).ToString
                Me.TextBox6.Text = DE
                Me.TextBox7.Text = SE
                DATO_EN_RRHH = True
            Else
                Me.TextBox4.Text = ""
                Me.TextBox6.Text = ""
                Me.TextBox7.Text = ""
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim DATO_1 As Boolean
    Private Sub BUSCAR_DATO_1()
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
                id_EMPLEADO = MiDataTable.Rows(0).Item(0).ToString
                Me.TextBox1.Text = MiDataTable.Rows(0).Item(3).ToString & " " & MiDataTable.Rows(0).Item(2).ToString
                Me.TextBox2.Text = MiDataTable.Rows(0).Item(1).ToString
                Me.TextBox8.Text = MiDataTable.Rows(0).Item(4).ToString
                Me.TextBox3.Text = Mid(MiDataTable.Rows(0).Item(5).ToString, 1, 10)
                Me.TextBox5.Text = MiDataTable.Rows(0).Item(8).ToString
                Me.TextBox36.Text = id_EMPLEADO
                DATO_1 = True
            Else
                id_EMPLEADO = 0
                Me.TextBox1.Text = ""
                Me.TextBox2.Text = ""
                Me.TextBox3.Text = ""
                Me.TextBox8.Text = ""
                Me.TextBox4.Text = ""
                Me.TextBox5.Text = ""
                Me.TextBox36.Text = "0"
                DATO_1 = False
            End If
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
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE HIGIENE DOSIMETRIA]")
        DGV.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE HIGIENE DOSIMETRIA]")
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

        Me.DGV.Columns(1).HeaderText = "Departamento"
        Me.DGV.Columns(1).Width = 190
        Me.DGV.Columns(1).Visible = True
        Me.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(2).HeaderText = "Servicio/ Area"
        Me.DGV.Columns(2).Width = 180
        Me.DGV.Columns(2).Visible = True
        Me.DGV.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(3).HeaderText = "Cargo"
        Me.DGV.Columns(3).Width = 100
        Me.DGV.Columns(3).Visible = True
        Me.DGV.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(4).HeaderText = "ID_M_P"
        Me.DGV.Columns(4).Width = 0
        Me.DGV.Columns(4).Visible = False

        Me.DGV.Columns(5).HeaderText = "Codigo Dosimetro"
        Me.DGV.Columns(5).Width = 70
        Me.DGV.Columns(5).Visible = True
        Me.DGV.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(6).HeaderText = "Fecha Lectura"
        Me.DGV.Columns(6).Width = 70
        Me.DGV.Columns(6).Visible = True
        Me.DGV.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Me.DGV.Columns(7).HeaderText = "FECHA_PE_I"
        Me.DGV.Columns(7).Width = 0
        Me.DGV.Columns(7).Visible = False
        Me.DGV.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'me.DGV.Columns(7).DefaultCellStyle.BackColor = Color.Khaki

        Me.DGV.Columns(8).HeaderText = "FECHA_PE_F"
        Me.DGV.Columns(8).Width = 0
        Me.DGV.Columns(8).Visible = False
        Me.DGV.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'me.DGV.Columns(8).DefaultCellStyle.BackColor = Color.Khaki

        Me.DGV.Columns(9).HeaderText = "Dosis Eq. Per."
        Me.DGV.Columns(9).Width = 90
        Me.DGV.Columns(9).Visible = True
        Me.DGV.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'me.DGV.Columns(9).DefaultCellStyle.BackColor = Color.Khaki

        Me.DGV.Columns(10).HeaderText = "Dosis Acum. I S"
        Me.DGV.Columns(10).Width = 90
        Me.DGV.Columns(10).Visible = True
        Me.DGV.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(11).HeaderText = "Dosis Acum. II S"
        Me.DGV.Columns(11).Width = 90
        Me.DGV.Columns(11).Visible = True
        Me.DGV.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(12).HeaderText = "Dosis Acum. Anual"
        Me.DGV.Columns(12).Width = 90
        Me.DGV.Columns(12).Visible = True
        Me.DGV.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'me.DGV.Columns(12).DefaultCellStyle.BackColor = Color.Khaki

        Me.DGV.Columns(13).HeaderText = ""
        Me.DGV.Columns(13).Width = 0
        Me.DGV.Columns(13).Visible = False

        Me.DGV.Columns(14).HeaderText = ""
        Me.DGV.Columns(14).Width = 0
        Me.DGV.Columns(14).Visible = False

        Me.DGV.Columns(15).HeaderText = ""
        Me.DGV.Columns(15).Width = 0
        Me.DGV.Columns(15).Visible = False

        Me.DGV.Columns(16).HeaderText = "Periodo"
        Me.DGV.Columns(16).Width = 70
        Me.DGV.Columns(16).Visible = True
        Me.DGV.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(16).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'me.DGV.Columns(16).DefaultCellStyle.BackColor = Color.Khaki

        Me.DGV.Columns(17).HeaderText = "Observaciones"
        Me.DGV.Columns(17).Width = 0
        Me.DGV.Columns(17).Visible = False
        Me.DGV.Columns(17).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(17).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'me.DGV.Columns(17).DefaultCellStyle.BackColor = Color.Khaki
    End Sub
    Private Sub Frm_Dosimetria_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        HIGIENE_2 = False
    End Sub
    Private Sub Frm_Dosimetria_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        HIGIENE_2 = False
    End Sub
    Private Sub Frm_Dosimetria_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        HIGIENE_2 = False
    End Sub
    Private Sub Frm_Dosimetria_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        HIGIENE_2 = False
    End Sub
    Private Sub DGV_Click(sender As Object, e As EventArgs) Handles DGV.Click
        If Me.DGV.RowCount > 0 Then
            For I = 0 To Me.DGV.RowCount - 1
                If DGV.Rows(I).Selected = True Then
                    h5ID_DOS = Me.DGV.Item(0, I).Value
                    h5DEPARTAMENTO = Me.DGV.Item(1, I).Value
                    h5SERVICIO = Me.DGV.Item(2, I).Value
                    h5CARGO = Me.DGV.Item(3, I).Value
                    h5ID_M_P = Me.DGV.Item(4, I).Value
                    h5CODIGO_DOSIMETRO = Me.DGV.Item(5, I).Value
                    h5FECHA_LECTURA = Me.DGV.Item(6, I).Value
                    h5FECHA_PE_I = Me.DGV.Item(7, I).Value
                    h5FECHA_PE_F = Me.DGV.Item(8, I).Value
                    h5DOSIS_EP = Me.DGV.Item(9, I).Value
                    h5DOSIS_APS = Me.DGV.Item(10, I).Value
                    h5DOSIS_ASS = Me.DGV.Item(11, I).Value
                    h5DOSIS_AA = Me.DGV.Item(12, I).Value
                    h5FECHA_I_PORTADOR = Me.DGV.Item(13, I).Value
                    h5FECHA_C_PROGRAMADA = Me.DGV.Item(14, I).Value
                    h5ID_PER_DOSIME = Me.DGV.Item(15, I).Value
                    h5D_PER_DOSIME = Me.DGV.Item(16, I).Value
                    h5OBSERVACION = Me.DGV.Item(17, I).Value
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        Dim I As Integer
        For I = 0 To Me.DGV.RowCount - 1
            If Me.DGV.Item(0, I).Value = h5ID_DOS Then
                Me.DGV.Rows(I).Selected = True
                Me.DGV.CurrentCell = Me.DGV.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call VERIFICAR_ACCESOS("208") : If HAY_ACCESO = False Then : Exit Sub : End If
        If Val(h5ID_DOS) = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Dosimetria_Upd.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "Select * from [dbo].[VISTA MAESTRO DE HIGIENE DOSIMETRIA] WHERE ID_M_P = " & Me.TextBox36.Text & " ORDER BY FECHA_LECTURA"
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        End If
    End Sub
    Private Sub DGV_DoubleClick(sender As Object, e As EventArgs) Handles DGV.DoubleClick
        Call VERIFICAR_ACCESOS("208") : If HAY_ACCESO = False Then : Exit Sub : End If
        If Val(h5ID_DOS) = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Dosimetria_Upd.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "Select * from [dbo].[VISTA MAESTRO DE HIGIENE DOSIMETRIA] WHERE ID_M_P = " & Me.TextBox36.Text & " ORDER BY FECHA_LECTURA"
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call VERIFICAR_ACCESOS("207") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Dosimetria_Add.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "Select * from [dbo].[VISTA MAESTRO DE HIGIENE DOSIMETRIA] WHERE ID_M_P = " & Me.TextBox36.Text & " ORDER BY FECHA_LECTURA"
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call VERIFICAR_ACCESOS("209") : If HAY_ACCESO = False Then : Exit Sub : End If
        If Val(h5ID_DOS) = 0 Then
            MsgBox("Debe seleccionar el registro que desea Eliminar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Esta seguro de Eliminar el registro seleccionado?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            Call BORRAR_REGISTRO()
            CADENAsql = "Select * from [dbo].[VISTA MAESTRO DE HIGIENE DOSIMETRIA] WHERE ID_M_P = " & Me.TextBox36.Text & " ORDER BY FECHA_LECTURA"
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            h5ID_DOS = 0
        End If
    End Sub
    Private Sub BORRAR_REGISTRO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE HIGIENE DOSIMETRIA] WHERE ID_DOS = " & CInt(h5ID_DOS) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Call VERIFICAR_ACCESOS("210") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Dosimetria_Imp.ShowDialog()
    End Sub
End Class