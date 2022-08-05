Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Public Class Frm_Ex_Med_Ocup
    Private Sub Frm_Ex_Med_Ocup_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            HIGIENE_3 = False
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
        End If
    End Sub
    Private Sub Frm_Ex_Med_Ocup_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Frm_Principal.Timer1.Stop()
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Me.TextBox6.Text = ""
        Me.TextBox36.Text = "0"
        Me.TextBox8.Text = ""
        CADENAsql = "Select * from [dbo].[VISTA MAESTRO DE HIGIENE EXAMEN MEDICO OCUPACIONAL] WHERE ID_M_P = " & Me.TextBox36.Text & " ORDER BY FECHA_EVA"
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
        Me.TextBox2.Focus()
    End Sub
    Private Sub Button7_Click(sender As Object, e As System.EventArgs) Handles Button7.Click
        HIGIENE_3 = False
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
        Call VERIFICAR_ACCESOS("216") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Accidentes_Laborales_Buscar.ShowDialog()
        CE = h1CODIGO
        If CE <> "" Then
            Me.TextBox2.Text = CE
            vID_M_P = h1ID_M_P
            Me.TextBox36.Text = vID_M_P
            CADENAsql = "SELECT ID_M_P, CODIGO, APELLIDOS, NOMBRES, N_CEDULA, FEC_ING_PAME, ID_ESTADO_P, ID_N_PROFESIONAL, D_T_SEXO FROM [dbo].[VISTA MAESTRO DE PERSONAS] WHERE CODIGO = '" & Me.TextBox2.Text & "'"
            Call BUSCAR_DATO_1()
            If DATO_1 = True Then
                CADENAsql = "Select * from [dbo].[VISTA MAESTRO DE HIGIENE EXAMEN MEDICO OCUPACIONAL] WHERE ID_M_P = " & Me.TextBox36.Text & " ORDER BY FECHA_EVA"
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
            Call VERIFICAR_ACCESOS("216") : If HAY_ACCESO = False Then : Exit Sub : End If
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
                    CADENAsql = "Select * from [dbo].[VISTA MAESTRO DE HIGIENE EXAMEN MEDICO OCUPACIONAL] WHERE ID_M_P = " & Me.TextBox36.Text & " ORDER BY FECHA_EVA"
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
    Private Sub ARMAR_DATAGRIDVIEW()
        Me.DGV.Columns(0).HeaderText = "Id"
        Me.DGV.Columns(0).Width = 30
        Me.DGV.Columns(0).Visible = True
        Me.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(1).HeaderText = "Departamento"
        Me.DGV.Columns(1).Width = 170
        Me.DGV.Columns(1).Visible = True
        Me.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(2).HeaderText = "Servicio/ Area"
        Me.DGV.Columns(2).Width = 170
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

        Me.DGV.Columns(5).HeaderText = "Fecha Evaluacion"
        Me.DGV.Columns(5).Width = 80
        Me.DGV.Columns(5).Visible = True
        Me.DGV.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(6).HeaderText = "No. Expediente"
        Me.DGV.Columns(6).Width = 70
        Me.DGV.Columns(6).Visible = True
        Me.DGV.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(7).HeaderText = "ID_T_EVALUACION"
        Me.DGV.Columns(7).Width = 0
        Me.DGV.Columns(7).Visible = False

        Me.DGV.Columns(8).HeaderText = "Tipo Evaluacion"
        Me.DGV.Columns(8).Width = 90
        Me.DGV.Columns(8).Visible = True
        Me.DGV.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(9).HeaderText = "Cargo Anterior"
        Me.DGV.Columns(9).Width = 90
        Me.DGV.Columns(9).Visible = True
        Me.DGV.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(10).HeaderText = "Cargo Actual"
        Me.DGV.Columns(10).Width = 90
        Me.DGV.Columns(10).Visible = True
        Me.DGV.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(11).HeaderText = "Exposicion Anterior"
        Me.DGV.Columns(11).Width = 90
        Me.DGV.Columns(11).Visible = True
        Me.DGV.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(12).HeaderText = "Exposicion Actual"
        Me.DGV.Columns(12).Width = 90
        Me.DGV.Columns(12).Visible = True
        Me.DGV.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(13).HeaderText = "Diagnostico"
        Me.DGV.Columns(13).Width = 0
        Me.DGV.Columns(13).Visible = False
        Me.DGV.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(14).HeaderText = "Recomendaciones"
        Me.DGV.Columns(14).Width = 0
        Me.DGV.Columns(14).Visible = False
        Me.DGV.Columns(14).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(14).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(15).HeaderText = "IR_R_EVALUACION"
        Me.DGV.Columns(15).Width = 0
        Me.DGV.Columns(15).Visible = False
        Me.DGV.Columns(15).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(15).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(16).HeaderText = "Resultado Evaluacion"
        Me.DGV.Columns(16).Width = 0
        Me.DGV.Columns(16).Visible = False
        Me.DGV.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(16).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(17).HeaderText = "OBSERVACIONES"
        Me.DGV.Columns(17).Width = 0
        Me.DGV.Columns(17).Visible = False
        Me.DGV.Columns(17).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(17).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

    End Sub
    Private Sub CARGAR_DATAGRIDVIEW()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE HIGIENE EXAMEN MEDICO OCUPACIONAL]")
        DGV.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE HIGIENE EXAMEN MEDICO OCUPACIONAL]")
        Me.Label2.Text = DGV.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub Frm_Dosimetria_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        HIGIENE_3 = False
    End Sub
    Private Sub Frm_Dosimetria_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        HIGIENE_3 = False
    End Sub
    Private Sub Frm_Dosimetria_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        HIGIENE_3 = False
    End Sub
    Private Sub Frm_Dosimetria_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        HIGIENE_3 = False
    End Sub
    Private Sub DGV_Click(sender As Object, e As EventArgs) Handles DGV.Click
        If Me.DGV.RowCount > 0 Then
            For I = 0 To Me.DGV.RowCount - 1
                If DGV.Rows(I).Selected = True Then
                    emoID_EMO = Me.DGV.Item(0, I).Value
                    emoDEPARTAMENTO = Me.DGV.Item(1, I).Value
                    emoSERVICIO = Me.DGV.Item(2, I).Value
                    emoCARGO = Me.DGV.Item(3, I).Value
                    emoID_M_P = Me.DGV.Item(4, I).Value
                    emoFECHA_EVA = Me.DGV.Item(5, I).Value
                    emoNO_EXPEDIENTE = Me.DGV.Item(6, I).Value
                    emoID_T_EVALUACION = Me.DGV.Item(7, I).Value
                    emoD_T_EVALUACION = Me.DGV.Item(8, I).Value
                    emoCARGO_ANTERIOR = Me.DGV.Item(9, I).Value
                    emoCARGO_ACTUAL = Me.DGV.Item(10, I).Value
                    emoAE_ANTERIOR = Me.DGV.Item(11, I).Value
                    emoAE_ACTUAL = Me.DGV.Item(12, I).Value
                    emoDIAGNOSTICO = Me.DGV.Item(13, I).Value
                    emoRECOMENDACIONES = Me.DGV.Item(14, I).Value
                    emoID_R_EVALUACION = Me.DGV.Item(15, I).Value
                    emoD_R_EVALUACION = Me.DGV.Item(16, I).Value
                    emoOBSERVACIONES = Me.DGV.Item(17, I).Value
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        Dim I As Integer
        For I = 0 To Me.DGV.RowCount - 1
            If Me.DGV.Item(0, I).Value = emoID_EMO Then
                Me.DGV.Rows(I).Selected = True
                Me.DGV.CurrentCell = Me.DGV.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call VERIFICAR_ACCESOS("212") : If HAY_ACCESO = False Then : Exit Sub : End If
        If TextBox36.Text = "" Or TextBox36.Text = "0" Then
            MsgBox("Debe seleccionar al Empleado correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox2.Focus()
            Exit Sub
        End If
        Frm_Ex_Med_Ocup_Add.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "Select * from [dbo].[VISTA MAESTRO DE HIGIENE EXAMEN MEDICO OCUPACIONAL] WHERE ID_M_P = " & Me.TextBox36.Text & " ORDER BY FECHA_EVA"
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call VERIFICAR_ACCESOS("214") : If HAY_ACCESO = False Then : Exit Sub : End If
        If Val(emoID_EMO) = 0 Then
            MsgBox("Debe seleccionar el registro que desea Eliminar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Esta seguro de Eliminar el registro seleccionado?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            Call BORRAR_REGISTRO_2()
            Call BORRAR_REGISTRO_1()
            CADENAsql = "Select * from [dbo].[VISTA MAESTRO DE HIGIENE EXAMEN MEDICO OCUPACIONAL] WHERE ID_M_P = " & Me.TextBox36.Text & " ORDER BY FECHA_EVA"
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            emoID_EMO = 0
        End If
    End Sub
    Private Sub BORRAR_REGISTRO_2()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE HIGIENE EXAMEN MEDICO OCUPACIONAL EXAMENES] WHERE ID_EMO = " & CInt(emoID_EMO) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub BORRAR_REGISTRO_1()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE HIGIENE EXAMEN MEDICO OCUPACIONAL] WHERE ID_EMO = " & CInt(emoID_EMO) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call VERIFICAR_ACCESOS("213") : If HAY_ACCESO = False Then : Exit Sub : End If
        If Val(emoID_EMO) = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Ex_Med_Ocup_Upd.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "Select * from [dbo].[VISTA MAESTRO DE HIGIENE EXAMEN MEDICO OCUPACIONAL] WHERE ID_M_P = " & Me.TextBox36.Text & " ORDER BY FECHA_EVA"
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        End If
    End Sub
    Private Sub DGV_DoubleClick(sender As Object, e As EventArgs) Handles DGV.DoubleClick
        Call VERIFICAR_ACCESOS("213") : If HAY_ACCESO = False Then : Exit Sub : End If
        If Val(emoID_EMO) = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Ex_Med_Ocup_Upd.ShowDialog()
        If CERRAR = False Then
            CADENAsql = "Select * from [dbo].[VISTA MAESTRO DE HIGIENE EXAMEN MEDICO OCUPACIONAL] WHERE ID_M_P = " & Me.TextBox36.Text & " ORDER BY FECHA_EVA"
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        End If
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Call VERIFICAR_ACCESOS("215") : If HAY_ACCESO = False Then : Exit Sub : End If
        If Val(Me.TextBox36.Text) = 0 Then
            MsgBox("Debe seleccionar el Empleado del que desea Imprimir los Examenes Medicos Ocupacionales", vbInformation, "Mensaje del Sistema")
            Me.TextBox2.Text = ""
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Desea Imprimir Todos los registros del Empleado?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            SELECCION = "({MAESTRO_DE_HIGIENE_EXAMEN_MEDICO_OCUPACIONAL.ID_M_P} = " & CInt(Me.TextBox36.Text) & ")"
        Else
            If emoID_EMO = 0 Then
                MsgBox("Debe seleccionar el registro que desea Imprimir", vbInformation, "Mensaje del Sistema")
                Me.DGV.Focus()
                Exit Sub
            Else
                SELECCION = "({MAESTRO_DE_HIGIENE_EXAMEN_MEDICO_OCUPACIONAL.ID_M_P} = " & CInt(Me.TextBox36.Text) & ") AND ({MAESTRO_DE_HIGIENE_EXAMEN_MEDICO_OCUPACIONAL.ID_EMO} = " & CInt(emoID_EMO) & ")"
            End If
        End If
        PARAMETRO = 19
        Frm_Visor_Reportes.CrystalR.ShowExportButton = True
        Frm_Visor_Reportes.CrystalR.ShowPrintButton = True
        Frm_Visor_Reportes.ShowDialog()
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Call VERIFICAR_ACCESOS("217") : If HAY_ACCESO = False Then : Exit Sub : End If
        If Val(emoID_EMO) = 0 Then
            MsgBox("Debe seleccionar el registro del que desea Visualizar los Examenes", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Examenes.ShowDialog()
    End Sub
End Class