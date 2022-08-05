Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Agregar_Subsidios_SS
    Dim ANTxDIA As Double
    Dim valTSit As Integer
    Dim valSit As Integer
    Dim valValor As String
    Private Sub Frm_Agregar_Subsidios_SS_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Agregar_Subsidios_SS_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW()
        If Conexion.CxSUBS.State = ConnectionState.Open Then
            Conexion.CBD_SUBS()
        End If
        Conexion.ABD_SUBS()
        Dim MiAdaptadorX As New SqlClient.SqlDataAdapter("SELECT * FROM [VISTA MAESTRO DE SUBSIDIOS] WHERE NO_INSS = '" & Frm_Expedientes_Activos.TextBox42.Text & "' ORDER BY FECHA_I", Conexion.CxSUBS)
        Dim MiDataSetX As New DataSet()
        MiAdaptadorX.Fill(MiDataSetX, "[VISTA MAESTRO DE SUBSIDIOS]")
        DGV1.DataSource = MiDataSetX.Tables("[VISTA MAESTRO DE SUBSIDIOS]")
        Me.Label103.Text = DGV1.RowCount & " Registro(s)"
        If Conexion.CxSUBS.State = ConnectionState.Open Then
            Conexion.CBD_SUBS()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW()
        Me.DGV1.Columns(0).HeaderText = "."
        Me.DGV1.Columns(0).Width = 10
        Me.DGV1.Columns(0).Visible = True
        Me.DGV1.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV1.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Me.DGV1.Columns(1).HeaderText = "No. Inss"
        Me.DGV1.Columns(1).Width = 0
        Me.DGV1.Columns(1).Visible = False
        Me.DGV1.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV1.Columns(1).DefaultCellStyle.BackColor = Color.Khaki

        Me.DGV1.Columns(2).HeaderText = "Afiliado"
        Me.DGV1.Columns(2).Width = 0
        Me.DGV1.Columns(2).Visible = False
        Me.DGV1.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(3).HeaderText = "ID_T_SEXO" : Me.DGV1.Columns(3).Width = 0 : Me.DGV1.Columns(3).Visible = False

        Me.DGV1.Columns(4).HeaderText = "Sexo"
        Me.DGV1.Columns(4).Width = 0
        Me.DGV1.Columns(4).Visible = False
        Me.DGV1.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(5).HeaderText = "ID_T_AFILIADO" : Me.DGV1.Columns(5).Width = 0 : Me.DGV1.Columns(5).Visible = False

        Me.DGV1.Columns(6).HeaderText = "Tipo Afiliado"
        Me.DGV1.Columns(6).Width = 0
        Me.DGV1.Columns(6).Visible = False
        Me.DGV1.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(7).HeaderText = "No. Expediente"
        Me.DGV1.Columns(7).Width = 50
        Me.DGV1.Columns(7).Visible = True
        Me.DGV1.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(8).HeaderText = "No. Boleta"
        Me.DGV1.Columns(8).Width = 50
        Me.DGV1.Columns(8).Visible = True
        Me.DGV1.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV1.Columns(8).DefaultCellStyle.BackColor = Color.LightSkyBlue

        Me.DGV1.Columns(9).HeaderText = "ID_T_SUBS" : Me.DGV1.Columns(9).Width = 0 : Me.DGV1.Columns(9).Visible = False

        Me.DGV1.Columns(10).HeaderText = "Tipo Subsidio"
        Me.DGV1.Columns(10).Width = 80
        Me.DGV1.Columns(10).Visible = True
        Me.DGV1.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(11).HeaderText = "No. Orden"
        Me.DGV1.Columns(11).Width = 40
        Me.DGV1.Columns(11).Visible = True
        Me.DGV1.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV1.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Me.DGV1.Columns(12).HeaderText = "Fec. Ini."
        Me.DGV1.Columns(12).Width = 70
        Me.DGV1.Columns(12).Visible = True
        Me.DGV1.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(13).HeaderText = "Fec. Fin."
        Me.DGV1.Columns(13).Width = 70
        Me.DGV1.Columns(13).Visible = True
        Me.DGV1.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(14).HeaderText = "Dias"
        Me.DGV1.Columns(14).Width = 30
        Me.DGV1.Columns(14).Visible = True
        Me.DGV1.Columns(14).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV1.Columns(14).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Me.DGV1.Columns(15).HeaderText = "Fec. Parto"
        Me.DGV1.Columns(15).Width = 70
        Me.DGV1.Columns(15).Visible = True
        Me.DGV1.Columns(15).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(16).HeaderText = "Fec. Posible Parto"
        Me.DGV1.Columns(16).Width = 70
        Me.DGV1.Columns(16).Visible = True
        Me.DGV1.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(17).HeaderText = "Fec. Accidente"
        Me.DGV1.Columns(17).Width = 70
        Me.DGV1.Columns(17).Visible = True
        Me.DGV1.Columns(17).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(18).HeaderText = "Fec. Declaracion"
        Me.DGV1.Columns(18).Width = 70
        Me.DGV1.Columns(18).Visible = True
        Me.DGV1.Columns(18).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(19).HeaderText = "VALOR_X_DIA" : Me.DGV1.Columns(19).Width = 0 : Me.DGV1.Columns(19).Visible = False

        Me.DGV1.Columns(20).HeaderText = "VALOR_TOTAL" : Me.DGV1.Columns(20).Width = 0 : Me.DGV1.Columns(20).Visible = False

        Me.DGV1.Columns(21).HeaderText = "ID_P_MEDICO" : Me.DGV1.Columns(21).Width = 0 : Me.DGV1.Columns(21).Visible = False

        Me.DGV1.Columns(22).HeaderText = "Medico"
        Me.DGV1.Columns(22).Width = 100
        Me.DGV1.Columns(22).Visible = True
        Me.DGV1.Columns(22).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(23).HeaderText = "ID_TIPO_EM" : Me.DGV1.Columns(23).Width = 0 : Me.DGV1.Columns(23).Visible = False

        Me.DGV1.Columns(24).HeaderText = "Especialidad Medica"
        Me.DGV1.Columns(24).Width = 80
        Me.DGV1.Columns(24).Visible = True
        Me.DGV1.Columns(24).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(25).HeaderText = "Diagnostico"
        Me.DGV1.Columns(25).Width = 70
        Me.DGV1.Columns(25).Visible = True
        Me.DGV1.Columns(25).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(26).HeaderText = "ID_ORGEN" : Me.DGV1.Columns(26).Width = 0 : Me.DGV1.Columns(26).Visible = False

        Me.DGV1.Columns(27).HeaderText = "Origen Subsidio"
        Me.DGV1.Columns(27).Width = 70
        Me.DGV1.Columns(27).Visible = True
        Me.DGV1.Columns(27).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(28).HeaderText = "Observaciones"
        Me.DGV1.Columns(28).Width = 0
        Me.DGV1.Columns(28).Visible = False
        Me.DGV1.Columns(28).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(29).HeaderText = "Fec. Grabacion"
        Me.DGV1.Columns(29).Width = 0
        Me.DGV1.Columns(29).Visible = False
        Me.DGV1.Columns(29).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(30).HeaderText = "USUARIO_ACT" : Me.DGV1.Columns(30).Width = 0 : Me.DGV1.Columns(30).Visible = False

        Me.DGV1.Columns(31).HeaderText = "EQUIPO_ACT" : Me.DGV1.Columns(31).Width = 0 : Me.DGV1.Columns(31).Visible = False

        Me.DGV1.Columns(32).HeaderText = "HORA_ACT" : Me.DGV1.Columns(32).Width = 0 : Me.DGV1.Columns(32).Visible = False
    End Sub
    Dim ssID_SUBS As Integer
    Dim ssNO_INSS As String
    Dim ssAFILIADO As String
    Dim ssID_T_SEXO As Integer
    Dim ssD_T_SEXO As String
    Dim ssID_T_AFILIADO As Integer
    Dim ssD_T_AFILIADO As String
    Dim ssNO_EXPEDIENTE As String
    Dim ssNO_BOLETA As String
    Dim ssID_T_SUBS As Integer
    Dim ssD_T_SUBS As String
    Dim ssN_ORDEN As String
    Dim ssFECHA_I As String
    Dim ssFECHA_F As String
    Dim ssCANT_DIAS As Integer
    Dim ssFECHA_PARTO As String
    Dim ssFECHA_PARTO_PROB As String
    Dim ssFECHA_ACC_LAB As String
    Dim ssFECHA_DECLARACION As String
    Dim ssVALOR_X_DIA As String
    Dim ssVALOR_TOTAL As String
    Dim ssID_P_MEDICO As Integer
    Dim ssMEDICO As String
    Dim ssID_TIPO_EM As Integer
    Dim ssD_TIPO_EM As String
    Dim ssDIAGNOSTICO As String
    Dim ssID_ORIGEN As Integer
    Dim ssD_ORIGEN As String
    Dim ssOBSERVACIONES As String
    Dim ssFECHA_ACT As String
    Dim ssUSUARIO_ACT As String
    Dim ssEQUIPO_ACT As String
    Dim ssHORA_ACT As String

    Private Sub DGV1_Click(sender As Object, e As EventArgs) Handles DGV1.Click
        If Me.DGV1.RowCount > 0 Then
            For I = 0 To Me.DGV1.RowCount - 1
                If DGV1.Rows(I).Selected = True Then
                    ssID_SUBS = Me.DGV1.Item(0, I).Value
                    ssNO_INSS = Me.DGV1.Item(1, I).Value
                    ssAFILIADO = Me.DGV1.Item(2, I).Value
                    ssID_T_SEXO = Me.DGV1.Item(3, I).Value
                    ssD_T_SEXO = Me.DGV1.Item(4, I).Value
                    ssID_T_AFILIADO = Me.DGV1.Item(5, I).Value
                    ssD_T_AFILIADO = Me.DGV1.Item(6, I).Value
                    ssNO_EXPEDIENTE = Me.DGV1.Item(7, I).Value
                    ssNO_BOLETA = Me.DGV1.Item(8, I).Value
                    ssID_T_SUBS = Me.DGV1.Item(9, I).Value
                    ssD_T_SUBS = Me.DGV1.Item(10, I).Value
                    ssN_ORDEN = Me.DGV1.Item(11, I).Value
                    ssFECHA_I = Me.DGV1.Item(12, I).Value
                    Me.DateTimePicker1.Value = ssFECHA_I
                    Me.DateTimePicker0.Value = ssFECHA_I
                    ssFECHA_F = Me.DGV1.Item(13, I).Value
                    Me.DateTimePicker2.Value = ssFECHA_F
                    ssCANT_DIAS = Me.DGV1.Item(14, I).Value
                    If Not IsDBNull(Me.DGV1.Item(15, I).Value) Then
                        ssFECHA_PARTO = Me.DGV1.Item(15, I).Value
                    Else
                        ssFECHA_PARTO = ""
                    End If
                    If Not IsDBNull(Me.DGV1.Item(16, I).Value) Then
                        ssFECHA_PARTO_PROB = Me.DGV1.Item(16, I).Value
                    Else
                        ssFECHA_PARTO_PROB = ""
                    End If
                    If Not IsDBNull(Me.DGV1.Item(17, I).Value) Then
                        ssFECHA_ACC_LAB = Me.DGV1.Item(17, I).Value
                    Else
                        ssFECHA_ACC_LAB = ""
                    End If
                    If Not IsDBNull(Me.DGV1.Item(18, I).Value) Then
                        ssFECHA_DECLARACION = Me.DGV1.Item(18, I).Value
                    Else
                        ssFECHA_DECLARACION = ""
                    End If
                    ssVALOR_X_DIA = Me.DGV1.Item(19, I).Value
                    ssVALOR_TOTAL = Me.DGV1.Item(20, I).Value
                    ssID_P_MEDICO = Me.DGV1.Item(21, I).Value
                    ssMEDICO = Me.DGV1.Item(22, I).Value
                    ssID_TIPO_EM = Me.DGV1.Item(23, I).Value
                    ssD_TIPO_EM = Me.DGV1.Item(24, I).Value
                    ssDIAGNOSTICO = Me.DGV1.Item(25, I).Value
                    ssID_ORIGEN = Me.DGV1.Item(26, I).Value
                    ssD_ORIGEN = Me.DGV1.Item(27, I).Value
                    ssOBSERVACIONES = Me.DGV1.Item(28, I).Value
                    ssFECHA_ACT = Me.DGV1.Item(29, I).Value
                    ssUSUARIO_ACT = Me.DGV1.Item(30, I).Value
                    ssEQUIPO_ACT = Me.DGV1.Item(31, I).Value
                    ssHORA_ACT = Me.DGV1.Item(32, I).Value
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub DGV1_DoubleClick(sender As Object, e As EventArgs) Handles DGV1.DoubleClick
        If Len(ssNO_BOLETA) = 0 Then
            MsgBox("Debe seleccionar el Documento que desea Trasladar", vbInformation, "Mensaje del Sistema")
            Me.DGV1.Focus()
            Exit Sub
        End If
        Call VERIFICAR_SI_YA_EXISTE_EL_DOCUMENTO()
        If YA_EXISTE_EL_DOCUMENTO = True Then
            MsgBox("El Documento que intenta trasladar ya existe en la Lista de Subsidios para esta Persona", vbInformation, "Mensaje del Sistema")
            Me.DGV1.Focus()
            Exit Sub
        End If
        MENSAJE = MsgBox("Este Proceso afectara las situaciones del Personal, ¿Esta seguro que desea continuar?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbNo Then
            Exit Sub
        End If
        Call GENERAR_ID()
        Call GRABAR_REGISTRO()

        vDET1 = "EXPEDIENTE: " & ssNO_EXPEDIENTE & "; NO. BOLETA: " & ssNO_EXPEDIENTE & "; FECHAS: " & Mid(ssFECHA_I, 1, 10) & " - " & Mid(ssFECHA_F, 1, 10)
        vDET2 = "MEDICO: (0) " & ssMEDICO
        vDET3 = "DIAGNOSTICO: " & ssDIAGNOSTICO

        Call GRABAR_SITUACION()
        vmssID_SUBS = ssID_SUBS
        Call CARGAR_DATAGRIDVIEW_1()
        Call ARMAR_DATAGRIDVIEW_1()
        Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_1()
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_1()
        Dim I As Integer
        For I = 0 To Frm_Expedientes_Activos.DGV17.RowCount - 1
            If Frm_Expedientes_Activos.DGV17.Item(0, I).Value = vmssID_SUBS Then
                Frm_Expedientes_Activos.DGV17.Rows(I).Selected = True
                Frm_Expedientes_Activos.DGV17.CurrentCell = Frm_Expedientes_Activos.DGV17.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_1()
        If Conexion.CxSUBS.State = ConnectionState.Open Then
            Conexion.CBD_SUBS()
        End If
        Conexion.ABD_SUBS()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxSUBS)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE SUBSIDIOS]")
        Frm_Expedientes_Activos.DGV17.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE SUBSIDIOS]")
        Frm_Expedientes_Activos.Label103.Text = "Total de Registros: " & Frm_Expedientes_Activos.DGV17.RowCount
        If Conexion.CxSUBS.State = ConnectionState.Open Then
            Conexion.CBD_SUBS()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_1()
        Frm_Expedientes_Activos.DGV17.Columns(0).HeaderText = "Id"
        Frm_Expedientes_Activos.DGV17.Columns(0).Width = 30
        Frm_Expedientes_Activos.DGV17.Columns(0).Visible = True
        Frm_Expedientes_Activos.DGV17.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV17.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Frm_Expedientes_Activos.DGV17.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Frm_Expedientes_Activos.DGV17.Columns(1).HeaderText = "ID_M_P"
        Frm_Expedientes_Activos.DGV17.Columns(1).Width = 0
        Frm_Expedientes_Activos.DGV17.Columns(1).Visible = False

        Frm_Expedientes_Activos.DGV17.Columns(2).HeaderText = "No. Exp."
        Frm_Expedientes_Activos.DGV17.Columns(2).Width = 80
        Frm_Expedientes_Activos.DGV17.Columns(2).Visible = True
        Frm_Expedientes_Activos.DGV17.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV17.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV17.Columns(3).HeaderText = "Boleta"
        Frm_Expedientes_Activos.DGV17.Columns(3).Width = 80
        Frm_Expedientes_Activos.DGV17.Columns(3).Visible = True
        Frm_Expedientes_Activos.DGV17.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV17.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV17.Columns(3).Frozen = True

        Frm_Expedientes_Activos.DGV17.Columns(4).HeaderText = "ID_T_SUBS"
        Frm_Expedientes_Activos.DGV17.Columns(4).Width = 0
        Frm_Expedientes_Activos.DGV17.Columns(4).Visible = False

        Frm_Expedientes_Activos.DGV17.Columns(5).HeaderText = "Tipo"
        Frm_Expedientes_Activos.DGV17.Columns(5).Width = 220
        Frm_Expedientes_Activos.DGV17.Columns(5).Visible = True
        Frm_Expedientes_Activos.DGV17.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV17.Columns(6).HeaderText = "Ord"
        Frm_Expedientes_Activos.DGV17.Columns(6).Width = 40
        Frm_Expedientes_Activos.DGV17.Columns(6).Visible = True
        Frm_Expedientes_Activos.DGV17.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV17.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Frm_Expedientes_Activos.DGV17.Columns(7).HeaderText = "Fecha Inicio"
        Frm_Expedientes_Activos.DGV17.Columns(7).Width = 90
        Frm_Expedientes_Activos.DGV17.Columns(7).Visible = True
        Frm_Expedientes_Activos.DGV17.Columns(7).DefaultCellStyle.Format = "d"
        Frm_Expedientes_Activos.DGV17.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV17.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV17.Columns(8).HeaderText = "Fecha Fin"
        Frm_Expedientes_Activos.DGV17.Columns(8).Width = 90
        Frm_Expedientes_Activos.DGV17.Columns(8).Visible = True
        Frm_Expedientes_Activos.DGV17.Columns(8).DefaultCellStyle.Format = "d"
        Frm_Expedientes_Activos.DGV17.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV17.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV17.Columns(9).HeaderText = "Dias"
        Frm_Expedientes_Activos.DGV17.Columns(9).Width = 40
        Frm_Expedientes_Activos.DGV17.Columns(9).Visible = True
        Frm_Expedientes_Activos.DGV17.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV17.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Frm_Expedientes_Activos.DGV17.Columns(10).HeaderText = "Fecha Parto"
        Frm_Expedientes_Activos.DGV17.Columns(10).Width = 90
        Frm_Expedientes_Activos.DGV17.Columns(10).Visible = True
        Frm_Expedientes_Activos.DGV17.Columns(10).DefaultCellStyle.Format = "d"
        Frm_Expedientes_Activos.DGV17.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV17.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV17.Columns(11).HeaderText = "Fecha Parto Prob."
        Frm_Expedientes_Activos.DGV17.Columns(11).Width = 90
        Frm_Expedientes_Activos.DGV17.Columns(11).Visible = True
        Frm_Expedientes_Activos.DGV17.Columns(11).DefaultCellStyle.Format = "d"
        Frm_Expedientes_Activos.DGV17.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV17.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV17.Columns(12).HeaderText = "Fecha Acc. Lab."
        Frm_Expedientes_Activos.DGV17.Columns(12).Width = 90
        Frm_Expedientes_Activos.DGV17.Columns(12).Visible = True
        Frm_Expedientes_Activos.DGV17.Columns(12).DefaultCellStyle.Format = "d"
        Frm_Expedientes_Activos.DGV17.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV17.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV17.Columns(13).HeaderText = "Fecha Declaracion"
        Frm_Expedientes_Activos.DGV17.Columns(13).Width = 90
        Frm_Expedientes_Activos.DGV17.Columns(13).Visible = True
        Frm_Expedientes_Activos.DGV17.Columns(13).DefaultCellStyle.Format = "d"
        Frm_Expedientes_Activos.DGV17.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV17.Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV17.Columns(14).HeaderText = "VALOR_X_DIA"
        Frm_Expedientes_Activos.DGV17.Columns(14).Width = 0
        Frm_Expedientes_Activos.DGV17.Columns(14).Visible = False

        Frm_Expedientes_Activos.DGV17.Columns(15).HeaderText = "VALOR_X_DIA"
        Frm_Expedientes_Activos.DGV17.Columns(15).Width = 0
        Frm_Expedientes_Activos.DGV17.Columns(15).Visible = False

        Frm_Expedientes_Activos.DGV17.Columns(16).HeaderText = "Medico"
        Frm_Expedientes_Activos.DGV17.Columns(16).Width = 230
        Frm_Expedientes_Activos.DGV17.Columns(16).Visible = True
        Frm_Expedientes_Activos.DGV17.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV17.Columns(17).HeaderText = "No. MINSA"
        Frm_Expedientes_Activos.DGV17.Columns(17).Width = 60
        Frm_Expedientes_Activos.DGV17.Columns(17).Visible = True
        Frm_Expedientes_Activos.DGV17.Columns(17).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV17.Columns(17).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV17.Columns(18).HeaderText = "ID_E_MED"
        Frm_Expedientes_Activos.DGV17.Columns(18).Width = 0
        Frm_Expedientes_Activos.DGV17.Columns(18).Visible = False

        Frm_Expedientes_Activos.DGV17.Columns(19).HeaderText = "Especialidad"
        Frm_Expedientes_Activos.DGV17.Columns(19).Width = 0
        Frm_Expedientes_Activos.DGV17.Columns(19).Visible = False
        Frm_Expedientes_Activos.DGV17.Columns(19).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV17.Columns(20).HeaderText = "Diagnostico"
        Frm_Expedientes_Activos.DGV17.Columns(20).Width = 500
        Frm_Expedientes_Activos.DGV17.Columns(20).Visible = True
        Frm_Expedientes_Activos.DGV17.Columns(20).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV17.Columns(21).HeaderText = "Observaciones"
        Frm_Expedientes_Activos.DGV17.Columns(21).Width = 100
        Frm_Expedientes_Activos.DGV17.Columns(21).Visible = True
        Frm_Expedientes_Activos.DGV17.Columns(21).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub GRABAR_REGISTRO()
        Dim fFECHA_I As String
        If ssFECHA_I <> "" Then
            fFECHA_I = Mid(ssFECHA_I, 1, 10)
        Else
            fFECHA_I = "NULL"
        End If
        Dim f1 As Object = If(fFECHA_I <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_I + "', 105), 23)", "NULL")

        Dim fFECHA_F As String
        If ssFECHA_F <> "" Then
            fFECHA_F = Mid(ssFECHA_F, 1, 10)
        Else
            fFECHA_F = "NULL"
        End If
        Dim f2 As Object = If(fFECHA_F <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_F + "', 105), 23)", "NULL")

        Dim fFECHA_PARTO As String
        If ssFECHA_PARTO <> "" Then
            fFECHA_PARTO = Mid(ssFECHA_PARTO, 1, 10)
        Else
            fFECHA_PARTO = "NULL"
        End If
        Dim f3 As Object = If(fFECHA_PARTO <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_PARTO + "', 105), 23)", "NULL")

        Dim fFECHA_PARTO_PROB As String
        If ssFECHA_PARTO_PROB <> "" Then
            fFECHA_PARTO_PROB = Mid(ssFECHA_PARTO_PROB, 1, 10)
        Else
            fFECHA_PARTO_PROB = "NULL"
        End If
        Dim f4 As Object = If(fFECHA_PARTO_PROB <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_PARTO_PROB + "', 105), 23)", "NULL")

        Dim fFECHA_ACC_LAB As String
        If ssFECHA_ACC_LAB <> "" Then
            fFECHA_ACC_LAB = Mid(ssFECHA_ACC_LAB, 1, 10)
        Else
            fFECHA_ACC_LAB = "NULL"
        End If
        Dim f5 As Object = If(fFECHA_ACC_LAB <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_ACC_LAB + "', 105), 23)", "NULL")

        Dim fFECHA_DECLARACION As String
        If ssFECHA_DECLARACION <> "" Then
            fFECHA_DECLARACION = Mid(ssFECHA_DECLARACION, 1, 10)
        Else
            fFECHA_DECLARACION = "NULL"
        End If
        Dim f6 As Object = If(fFECHA_DECLARACION <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_DECLARACION + "', 105), 23)", "NULL")

        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()

        Dim SQL As String
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            SQL = "INSERT INTO [MAESTRO DE SUBSIDIOS] (ID_SUBS, ID_M_P, NO_EXPEDIENTE, NO_BOLETA, ID_T_SUBS, N_ORDEN, FECHA_I, FECHA_F, CANT_DIAS, FECHA_PARTO, FECHA_PARTO_PROB, FECHA_ACC_LAB, FECHA_DECLARACION, VALOR_X_DIA, VALOR_TOTAL, MEDICO, N_MINSA, ID_E_MED, DIAGNOSTICO, OBSERVACIONES)" &
                   "values (" & idID_SUBS & ", " & Frm_Expedientes_Activos.TextBox35.Text & ", '" & ssNO_EXPEDIENTE & "','" & ssNO_BOLETA & "', " & ssID_T_SUBS & ", '" & ssN_ORDEN & "', " & f1 & ", " & f2 & ", " & ssCANT_DIAS & ", " & f3 & ", " & f4 & ", " & f5 & ", " & f6 & ", '" & ssVALOR_X_DIA & "', '" & ssVALOR_TOTAL & "', '" & ssMEDICO & "', '0', " & ssID_TIPO_EM & ", '" & ssDIAGNOSTICO & "', '" & ssOBSERVACIONES & "')"
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
    Dim idID_SUBS As Integer
    Private Sub GENERAR_ID()
        Dim CODIGO As Integer
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT ID_SUBS FROM [MAESTRO DE SUBSIDIOS] ORDER BY ID_SUBS", Conexion.CxRRHH)
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
            idID_SUBS = CODIGO
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim YA_EXISTE_EL_DOCUMENTO As Boolean
    Private Sub VERIFICAR_SI_YA_EXISTE_EL_DOCUMENTO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [VISTA MAESTRO DE SUBSIDIOS] WHERE NO_BOLETA = '" & Trim(ssNO_BOLETA) & "' ORDER BY ID_SUBS", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                YA_EXISTE_EL_DOCUMENTO = True
            Else
                YA_EXISTE_EL_DOCUMENTO = False
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub

    '//////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub GRABAR_SITUACION()
        Me.Cursor = Cursors.WaitCursor
        Me.DateTimePicker0.Value = Me.DateTimePicker1.Value
        Application.DoEvents()
        For I2 As Integer = 1 To DateDiff(DateInterval.Day, Me.DateTimePicker1.Value, Me.DateTimePicker2.Value.AddDays(1))
            Me.DateTimePicker0.Refresh()
            '-------------- SITUACIONES
            Call BUSCAR_SI_EXISTE_DIA_EN_MAESTRO_DE_SITUACION_DE_PERSONAL()
            If EXISTE_DIA_EN_MAESTRO_DE_SITUACION_DE_PERSONAL = True Then
                Call UPD_ROW()
            Else
                Call ADD_ROW()
            End If
            '-------------- SUBSIDIOS
            Call BUSCAR_SI_EXISTE_DIA_EN_MAESTRO_DE_SUBSIDIOS_DETALLES_DIAS()
            If EXISTE_DIA_EN_MAESTRO_DE_SUBSIDIOS_DETALLES_DIAS = True Then
                Call UPD_ROW_S()
            Else
                Call ADD_ROW_S()
            End If
            Me.DateTimePicker0.Value = Me.DateTimePicker0.Value.AddDays(1)
            Me.DateTimePicker0.Refresh()
        Next
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub UPD_ROW_S()

    End Sub
    Private Sub UPD_ROW()
        Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker0.Value, 1, 10) + "', 105), 23)"
        Dim CD As Integer = 30
        If CD = 30 Then
            ANTxDIA = 0.0833333333333333
        End If
        valTSit = 2
        valSit = 14
        valValor = 0
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE SITUACION DE PERSONAL] SET ID_T_SIT_P = " & valTSit & ", ID_SIT_P = " & valSit & ", VALOR_SIT = '" & CDbl(valValor) & "', VALOR_DIA = '" & CDbl(ANTxDIA) & "', OBSERVACIONES = 'INGRESADO DESDE EL MODULO DE EXPEDIENTES', DETALLE_1 = '" & vDET1 & "', DETALLE_2 = '" & vDET2 & "', DETALLE_3 = '" & vDET3 & "'  WHERE DIA = " & F1 & " AND ID_M_P = " & Frm_Expedientes_Activos.TextBox35.Text & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ADD_ROW()
        Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker0.Value, 1, 10) + "', 105), 23)"
        Dim CD As Integer = 30
        If CD = 30 Then
            ANTxDIA = 0.0833333333333333
        End If

        valTSit = 2
        valSit = 14
        valValor = 0
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE SITUACION DE PERSONAL] (ID_M_P, DIA, ID_T_SIT_P, ID_SIT_P, VALOR_SIT, VALOR_DIA, OBSERVACIONES, DETALLE_1, DETALLE_2, DETALLE_3)" &
                              "values (" & CInt(Frm_Expedientes_Activos.TextBox35.Text) & ", " & F1 & ", " & valTSit & ", " & valSit & ", '" & CDbl(valValor) & "', '" & CDbl(ANTxDIA) & "','INGRESADO DESDE EL MODULO DE EXPEDIENTES', '" & vDET1 & "', '" & vDET2 & "', '" & vDET3 & "')"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ADD_ROW_S()
        Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker0.Value, 1, 10) + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE SUBSIDIOS DETALLES DIAS] (ID_SUBS, CODIGO, NO_EXPEDIENTE, NO_BOLETA, ID_T_SUBS, FECHA_S, DIAGNOSTICO)" &
                              "values (" & idID_SUBS & ", '" & Frm_Expedientes_Activos.TextBox5.Text & "', '" & ssNO_EXPEDIENTE & "','" & ssNO_BOLETA & "', " & ssID_T_SUBS & ", " & F1 & ", '" & Trim(ssDIAGNOSTICO) & "')"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim EXISTE_DIA_EN_MAESTRO_DE_SITUACION_DE_PERSONAL As Boolean
    Private Sub BUSCAR_SI_EXISTE_DIA_EN_MAESTRO_DE_SITUACION_DE_PERSONAL()
        cEMPLEADO = Format(CInt(Frm_Expedientes_Activos.TextBox5.Text), "0000000000")
        Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker0.Value, 1, 10) + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable3 As New DataTable
        Dim MiDataAdapter3 As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter3 = New SqlDataAdapter("SELECT * FROM [VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE (CODIGO = '" & cEMPLEADO & "') AND (DIA = " & F1 & ") ORDER BY  CODIGO, DIA", Conexion.CxRRHH)
            MiDataTable3.Clear()
            MiDataAdapter3.Fill(MiDataTable3)
            If MiDataTable3.Rows.Count <> 0 Then
                EXISTE_DIA_EN_MAESTRO_DE_SITUACION_DE_PERSONAL = True
            Else
                EXISTE_DIA_EN_MAESTRO_DE_SITUACION_DE_PERSONAL = False
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim EXISTE_DIA_EN_MAESTRO_DE_SUBSIDIOS_DETALLES_DIAS As Boolean
    Private Sub BUSCAR_SI_EXISTE_DIA_EN_MAESTRO_DE_SUBSIDIOS_DETALLES_DIAS()
        cEMPLEADO = Format(CInt(Frm_Expedientes_Activos.TextBox5.Text), "0000000000")
        Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker0.Value, 1, 10) + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable3 As New DataTable
        Dim MiDataAdapter3 As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter3 = New SqlDataAdapter("SELECT * FROM [MAESTRO DE SUBSIDIOS DETALLES DIAS] WHERE (CODIGO = '" & cEMPLEADO & "') AND (FECHA_S = " & F1 & ") ORDER BY  CODIGO, FECHA_S", Conexion.CxRRHH)
            MiDataTable3.Clear()
            MiDataAdapter3.Fill(MiDataTable3)
            If MiDataTable3.Rows.Count <> 0 Then
                EXISTE_DIA_EN_MAESTRO_DE_SUBSIDIOS_DETALLES_DIAS = True
            Else
                EXISTE_DIA_EN_MAESTRO_DE_SUBSIDIOS_DETALLES_DIAS = False
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
End Class