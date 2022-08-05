Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Aspirantes
    Private Sub Frm_Aspirantes_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.CheckBox1.Checked = False
        If Me.CheckBox1.Checked = True Then
            Me.TextBox2.Enabled = True
        Else
            Me.TextBox2.Enabled = False
        End If
        Me.CheckBox2.Checked = False
        If Me.CheckBox2.Checked = True Then
            Me.DateTimePicker1.Enabled = True
            Me.DateTimePicker2.Enabled = True
            Me.Button10.Enabled = True
        Else
            Me.DateTimePicker1.Enabled = False
            Me.DateTimePicker2.Enabled = False
            Me.Button10.Enabled = False
        End If
        RemoveHandler Me.DGV1.SelectionChanged, AddressOf DGV1_SelectionChanged
        RemoveHandler Me.DGV2.RowPrePaint, AddressOf DGV2_RowPrePaint
        CADENAsql = "SELECT * FROM [VISTA SOL_MAESTRO DE SOLICITANTES] ORDER BY AN"
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
        CADENAsql = "SELECT * FROM [VISTA SOL_MAESTRO DE SOLICITANTES PROCESOS] WHERE ID_S = " & aspID_S & " ORDER BY FECHA_P"
        Call CARGAR_DATAGRIDVIEW_2()
        Call ARMAR_DATAGRIDVIEW_2()
        AddHandler Me.DGV2.RowPrePaint, AddressOf DGV2_RowPrePaint
        AddHandler Me.DGV1.SelectionChanged, AddressOf DGV1_SelectionChanged
    End Sub
    Private Sub Button7_Click(sender As Object, e As System.EventArgs) Handles Button7.Click
        ASPIRANTES = False
        Frm_Principal.Timer1.Start()
        Me.Close()
    End Sub
    Private Sub Frm_Aspirantes_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            ASPIRANTES = False
            Frm_Principal.Timer1.Start()
            Me.Close()
        End If
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA SOL_MAESTRO DE SOLICITANTES]")
        DGV1.DataSource = MiDataSet.Tables("[VISTA SOL_MAESTRO DE SOLICITANTES]")
        Me.Label2.Text = DGV1.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW()
        Me.DGV1.Columns(0).HeaderText = "Id"
        Me.DGV1.Columns(0).Width = 10
        Me.DGV1.Columns(0).Visible = True
        Me.DGV1.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV1.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Me.DGV1.Columns(1).HeaderText = "Cedula"
        Me.DGV1.Columns(1).Width = 100
        Me.DGV1.Columns(1).Visible = True
        Me.DGV1.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV1.Columns(1).DefaultCellStyle.BackColor = Color.LightBlue

        Me.DGV1.Columns(2).HeaderText = "APELLIDOS"
        Me.DGV1.Columns(2).Width = 0
        Me.DGV1.Columns(2).Visible = False

        Me.DGV1.Columns(3).HeaderText = "NOMBRES"
        Me.DGV1.Columns(3).Width = 0
        Me.DGV1.Columns(3).Visible = False

        Me.DGV1.Columns(4).HeaderText = "Aspirante"
        Me.DGV1.Columns(4).Width = 195
        Me.DGV1.Columns(4).Visible = True
        Me.DGV1.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(5).HeaderText = "Fecha Nacimiento"
        Me.DGV1.Columns(5).Width = 70
        Me.DGV1.Columns(5).Visible = True
        Me.DGV1.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV1.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV1.Columns(5).DefaultCellStyle.BackColor = Color.LightGreen

        Me.DGV1.Columns(6).HeaderText = ""
        Me.DGV1.Columns(6).Width = 15
        Me.DGV1.Columns(6).Visible = True
        Me.DGV1.Columns(6).DefaultCellStyle.BackColor = Color.Black
        Me.DGV1.Columns(6).Frozen = True

        Me.DGV1.Columns(7).HeaderText = "Perfil"
        Me.DGV1.Columns(7).Width = 100
        Me.DGV1.Columns(7).Visible = True
        Me.DGV1.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(8).HeaderText = "ID_T_SEXO"
        Me.DGV1.Columns(8).Width = 0
        Me.DGV1.Columns(8).Visible = False

        Me.DGV1.Columns(9).HeaderText = "Sexo"
        Me.DGV1.Columns(9).Width = 60
        Me.DGV1.Columns(9).Visible = True
        Me.DGV1.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(10).HeaderText = "Direccion 1"
        Me.DGV1.Columns(10).Width = 150
        Me.DGV1.Columns(10).Visible = True
        Me.DGV1.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(11).HeaderText = "Direccion 2"
        Me.DGV1.Columns(11).Width = 150
        Me.DGV1.Columns(11).Visible = True
        Me.DGV1.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(12).HeaderText = "Telefono Claro"
        Me.DGV1.Columns(12).Width = 70
        Me.DGV1.Columns(12).Visible = True
        Me.DGV1.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV1.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(13).HeaderText = "Telefono Tigo"
        Me.DGV1.Columns(13).Width = 70
        Me.DGV1.Columns(13).Visible = True
        Me.DGV1.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV1.Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(14).HeaderText = "Correo 1"
        Me.DGV1.Columns(14).Width = 100
        Me.DGV1.Columns(14).Visible = True
        Me.DGV1.Columns(14).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(15).HeaderText = "Correo 2"
        Me.DGV1.Columns(15).Width = 100
        Me.DGV1.Columns(15).Visible = True
        Me.DGV1.Columns(15).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(16).HeaderText = "Whatsapp"
        Me.DGV1.Columns(16).Width = 70
        Me.DGV1.Columns(16).Visible = True
        Me.DGV1.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV1.Columns(16).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(17).HeaderText = "Padre"
        Me.DGV1.Columns(17).Width = 70
        Me.DGV1.Columns(17).Visible = True
        Me.DGV1.Columns(17).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(18).HeaderText = "Madre"
        Me.DGV1.Columns(18).Width = 100
        Me.DGV1.Columns(18).Visible = True
        Me.DGV1.Columns(18).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(19).HeaderText = "Referencia Academica"
        Me.DGV1.Columns(19).Width = 150
        Me.DGV1.Columns(19).Visible = True
        Me.DGV1.Columns(19).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(20).HeaderText = "Referencia Estudios"
        Me.DGV1.Columns(20).Width = 150
        Me.DGV1.Columns(20).Visible = True
        Me.DGV1.Columns(20).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(21).HeaderText = "Recomendado por"
        Me.DGV1.Columns(21).Width = 150
        Me.DGV1.Columns(21).Visible = True
        Me.DGV1.Columns(21).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(22).HeaderText = "Observaciones"
        Me.DGV1.Columns(22).Width = 150
        Me.DGV1.Columns(22).Visible = True
        Me.DGV1.Columns(22).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV1.Columns(23).HeaderText = "ID_ESTADO_R"
        Me.DGV1.Columns(23).Width = 0
        Me.DGV1.Columns(23).Visible = False

        Me.DGV1.Columns(24).HeaderText = "Fecha Ingreso"
        Me.DGV1.Columns(24).Width = 70
        Me.DGV1.Columns(24).Visible = True
        Me.DGV1.Columns(24).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV1.Columns(24).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV1.Columns(24).DefaultCellStyle.BackColor = Color.Khaki
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_2()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA SOL_MAESTRO DE SOLICITANTES PROCESOS]")
        DGV2.DataSource = MiDataSet.Tables("[VISTA SOL_MAESTRO DE SOLICITANTES PROCESOS]")
        'Me.Label2.Text = DGV1.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_2()
        Me.DGV2.Columns(0).HeaderText = "Id"
        Me.DGV2.Columns(0).Width = 10
        Me.DGV2.Columns(0).Visible = True
        Me.DGV2.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV2.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV2.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Me.DGV2.Columns(1).HeaderText = "ID_S"
        Me.DGV2.Columns(1).Width = 0
        Me.DGV2.Columns(1).Visible = False

        Me.DGV2.Columns(2).HeaderText = "Fecha Proceso"
        Me.DGV2.Columns(2).Width = 70
        Me.DGV2.Columns(2).Visible = True
        Me.DGV2.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV2.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV2.Columns(3).HeaderText = "ID_PROC"
        Me.DGV2.Columns(3).Width = 0
        Me.DGV2.Columns(3).Visible = False

        Me.DGV2.Columns(4).HeaderText = "Proceso"
        Me.DGV2.Columns(4).Width = 120
        Me.DGV2.Columns(4).Visible = True
        Me.DGV2.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV2.Columns(4).Frozen = True

        Me.DGV2.Columns(5).HeaderText = "ID_U_RECURSO"
        Me.DGV2.Columns(5).Width = 0
        Me.DGV2.Columns(5).Visible = False

        Me.DGV2.Columns(6).HeaderText = "Ubicacion Recurso"
        Me.DGV2.Columns(6).Width = 80
        Me.DGV2.Columns(6).Visible = True
        Me.DGV2.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV2.Columns(7).HeaderText = "Reali zado"
        Me.DGV2.Columns(7).Width = 40
        Me.DGV2.Columns(7).Visible = True
        Me.DGV2.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV2.Columns(8).HeaderText = "ID_R_SOL"
        Me.DGV2.Columns(8).Width = 0
        Me.DGV2.Columns(8).Visible = False

        Me.DGV2.Columns(9).HeaderText = "Resultado"
        Me.DGV2.Columns(9).Width = 200
        Me.DGV2.Columns(9).Visible = True
        Me.DGV2.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV2.Columns(10).HeaderText = "Observaciones"
        Me.DGV2.Columns(10).Width = 540
        Me.DGV2.Columns(10).Visible = True
        Me.DGV2.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV2.Columns(11).HeaderText = "ID_ESTADO_R"
        Me.DGV2.Columns(11).Width = 0
        Me.DGV2.Columns(11).Visible = False
    End Sub
    Private Sub DGV1_Click(sender As Object, e As EventArgs) Handles DGV1.Click
        If DGV1.SelectedCells.Count <> 0 Then
            Dim FILA = DGV1.CurrentRow.Index
            If Not IsDBNull(Me.DGV1.Item(0, FILA).Value) Then : aspID_S = Me.DGV1.Item(0, FILA).Value : Else : aspID_S = 0 : End If
            If Not IsDBNull(Me.DGV1.Item(1, FILA).Value) Then : aspCEDULA = Me.DGV1.Item(1, FILA).Value : Else : aspCEDULA = "" : End If
            If Not IsDBNull(Me.DGV1.Item(2, FILA).Value) Then : aspAPELLIDOS = Me.DGV1.Item(2, FILA).Value : Else : aspAPELLIDOS = "" : End If
            If Not IsDBNull(Me.DGV1.Item(3, FILA).Value) Then : aspNOMBRES = Me.DGV1.Item(3, FILA).Value : Else : aspNOMBRES = "" : End If
            If Not IsDBNull(Me.DGV1.Item(4, FILA).Value) Then : aspAN = Me.DGV1.Item(4, FILA).Value : Else : aspAN = "" : End If
            If Not IsDBNull(Me.DGV1.Item(5, FILA).Value) Then : aspFECHA_NAC = Me.DGV1.Item(5, FILA).Value : Else : aspFECHA_NAC = "" : End If
            'SEPARADOR
            If Not IsDBNull(Me.DGV1.Item(7, FILA).Value) Then : aspPERFIL = Me.DGV1.Item(7, FILA).Value : Else : aspPERFIL = "" : End If
            If Not IsDBNull(Me.DGV1.Item(8, FILA).Value) Then : aspID_T_SEXO = Me.DGV1.Item(8, FILA).Value : Else : aspID_T_SEXO = 0 : End If
            If Not IsDBNull(Me.DGV1.Item(9, FILA).Value) Then : aspD_T_SEXO = Me.DGV1.Item(9, FILA).Value : Else : aspD_T_SEXO = "" : End If
            If Not IsDBNull(Me.DGV1.Item(10, FILA).Value) Then : aspDIRECCION_1 = Me.DGV1.Item(10, FILA).Value : Else : aspDIRECCION_1 = "" : End If
            If Not IsDBNull(Me.DGV1.Item(11, FILA).Value) Then : aspDIRECCION_2 = Me.DGV1.Item(11, FILA).Value : Else : aspDIRECCION_2 = "" : End If
            If Not IsDBNull(Me.DGV1.Item(12, FILA).Value) Then : aspTELEFONO_1 = Me.DGV1.Item(12, FILA).Value : Else : aspTELEFONO_1 = "" : End If
            If Not IsDBNull(Me.DGV1.Item(13, FILA).Value) Then : aspTELEFONO_2 = Me.DGV1.Item(13, FILA).Value : Else : aspTELEFONO_2 = "" : End If
            If Not IsDBNull(Me.DGV1.Item(14, FILA).Value) Then : aspCORREO_1 = Me.DGV1.Item(14, FILA).Value : Else : aspCORREO_1 = "" : End If
            If Not IsDBNull(Me.DGV1.Item(15, FILA).Value) Then : aspCORREO_2 = Me.DGV1.Item(15, FILA).Value : Else : aspCORREO_2 = "" : End If
            If Not IsDBNull(Me.DGV1.Item(16, FILA).Value) Then : aspWHATSAPP = Me.DGV1.Item(16, FILA).Value : Else : aspWHATSAPP = "" : End If
            If Not IsDBNull(Me.DGV1.Item(17, FILA).Value) Then : aspPADRE = Me.DGV1.Item(17, FILA).Value : Else : aspPADRE = "" : End If
            If Not IsDBNull(Me.DGV1.Item(18, FILA).Value) Then : aspMADRE = Me.DGV1.Item(18, FILA).Value : Else : aspMADRE = "" : End If
            If Not IsDBNull(Me.DGV1.Item(19, FILA).Value) Then : aspREFERENCIA_ACADEMICA = Me.DGV1.Item(19, FILA).Value : Else : aspREFERENCIA_ACADEMICA = "" : End If
            If Not IsDBNull(Me.DGV1.Item(20, FILA).Value) Then : aspREFERENCIA_ESTUDIOS = Me.DGV1.Item(20, FILA).Value : Else : aspREFERENCIA_ESTUDIOS = "" : End If
            If Not IsDBNull(Me.DGV1.Item(21, FILA).Value) Then : aspRECOMENDADO_POR = Me.DGV1.Item(21, FILA).Value : Else : aspRECOMENDADO_POR = "" : End If
            If Not IsDBNull(Me.DGV1.Item(22, FILA).Value) Then : aspOBSERVACIONES = Me.DGV1.Item(22, FILA).Value : Else : aspOBSERVACIONES = "" : End If
            If Not IsDBNull(Me.DGV1.Item(23, FILA).Value) Then : aspID_ESTADO_R = Me.DGV1.Item(23, FILA).Value : Else : aspID_ESTADO_R = 0 : End If
            If Not IsDBNull(Me.DGV1.Item(24, FILA).Value) Then : aspFECHA_I = Me.DGV1.Item(24, FILA).Value : Else : aspFECHA_I = "" : End If
            CADENAsql = "SELECT * FROM [VISTA SOL_MAESTRO DE SOLICITANTES PROCESOS] WHERE ID_S = " & aspID_S & " ORDER BY FECHA_P"
            Call CARGAR_DATAGRIDVIEW_2()
            Call ARMAR_DATAGRIDVIEW_2()
        End If
    End Sub
    Private Sub DGV1_SelectionChanged(sender As Object, e As EventArgs) Handles DGV1.SelectionChanged
        If DGV1.SelectedCells.Count <> 0 Then
            Dim FILA = DGV1.CurrentRow.Index
            If Not IsDBNull(Me.DGV1.Item(0, FILA).Value) Then : aspID_S = Me.DGV1.Item(0, FILA).Value : Else : aspID_S = 0 : End If
            If Not IsDBNull(Me.DGV1.Item(1, FILA).Value) Then : aspCEDULA = Me.DGV1.Item(1, FILA).Value : Else : aspCEDULA = "" : End If
            If Not IsDBNull(Me.DGV1.Item(2, FILA).Value) Then : aspAPELLIDOS = Me.DGV1.Item(2, FILA).Value : Else : aspAPELLIDOS = "" : End If
            If Not IsDBNull(Me.DGV1.Item(3, FILA).Value) Then : aspNOMBRES = Me.DGV1.Item(3, FILA).Value : Else : aspNOMBRES = "" : End If
            If Not IsDBNull(Me.DGV1.Item(4, FILA).Value) Then : aspAN = Me.DGV1.Item(4, FILA).Value : Else : aspAN = "" : End If
            If Not IsDBNull(Me.DGV1.Item(5, FILA).Value) Then : aspFECHA_NAC = Me.DGV1.Item(5, FILA).Value : Else : aspFECHA_NAC = "" : End If
            'SEPARADOR
            If Not IsDBNull(Me.DGV1.Item(7, FILA).Value) Then : aspPERFIL = Me.DGV1.Item(7, FILA).Value : Else : aspPERFIL = "" : End If
            If Not IsDBNull(Me.DGV1.Item(8, FILA).Value) Then : aspID_T_SEXO = Me.DGV1.Item(8, FILA).Value : Else : aspID_T_SEXO = 0 : End If
            If Not IsDBNull(Me.DGV1.Item(9, FILA).Value) Then : aspD_T_SEXO = Me.DGV1.Item(9, FILA).Value : Else : aspD_T_SEXO = "" : End If
            If Not IsDBNull(Me.DGV1.Item(10, FILA).Value) Then : aspDIRECCION_1 = Me.DGV1.Item(10, FILA).Value : Else : aspDIRECCION_1 = "" : End If
            If Not IsDBNull(Me.DGV1.Item(11, FILA).Value) Then : aspDIRECCION_2 = Me.DGV1.Item(11, FILA).Value : Else : aspDIRECCION_2 = "" : End If
            If Not IsDBNull(Me.DGV1.Item(12, FILA).Value) Then : aspTELEFONO_1 = Me.DGV1.Item(12, FILA).Value : Else : aspTELEFONO_1 = "" : End If
            If Not IsDBNull(Me.DGV1.Item(13, FILA).Value) Then : aspTELEFONO_2 = Me.DGV1.Item(13, FILA).Value : Else : aspTELEFONO_2 = "" : End If
            If Not IsDBNull(Me.DGV1.Item(14, FILA).Value) Then : aspCORREO_1 = Me.DGV1.Item(14, FILA).Value : Else : aspCORREO_1 = "" : End If
            If Not IsDBNull(Me.DGV1.Item(15, FILA).Value) Then : aspCORREO_2 = Me.DGV1.Item(15, FILA).Value : Else : aspCORREO_2 = "" : End If
            If Not IsDBNull(Me.DGV1.Item(16, FILA).Value) Then : aspWHATSAPP = Me.DGV1.Item(16, FILA).Value : Else : aspWHATSAPP = "" : End If
            If Not IsDBNull(Me.DGV1.Item(17, FILA).Value) Then : aspPADRE = Me.DGV1.Item(17, FILA).Value : Else : aspPADRE = "" : End If
            If Not IsDBNull(Me.DGV1.Item(18, FILA).Value) Then : aspMADRE = Me.DGV1.Item(18, FILA).Value : Else : aspMADRE = "" : End If
            If Not IsDBNull(Me.DGV1.Item(19, FILA).Value) Then : aspREFERENCIA_ACADEMICA = Me.DGV1.Item(19, FILA).Value : Else : aspREFERENCIA_ACADEMICA = "" : End If
            If Not IsDBNull(Me.DGV1.Item(20, FILA).Value) Then : aspREFERENCIA_ESTUDIOS = Me.DGV1.Item(20, FILA).Value : Else : aspREFERENCIA_ESTUDIOS = "" : End If
            If Not IsDBNull(Me.DGV1.Item(21, FILA).Value) Then : aspRECOMENDADO_POR = Me.DGV1.Item(21, FILA).Value : Else : aspRECOMENDADO_POR = "" : End If
            If Not IsDBNull(Me.DGV1.Item(22, FILA).Value) Then : aspOBSERVACIONES = Me.DGV1.Item(22, FILA).Value : Else : aspOBSERVACIONES = "" : End If
            If Not IsDBNull(Me.DGV1.Item(23, FILA).Value) Then : aspID_ESTADO_R = Me.DGV1.Item(23, FILA).Value : Else : aspID_ESTADO_R = 0 : End If
            If Not IsDBNull(Me.DGV1.Item(24, FILA).Value) Then : aspFECHA_I = Me.DGV1.Item(24, FILA).Value : Else : aspFECHA_I = "" : End If
            CADENAsql = "SELECT * FROM [VISTA SOL_MAESTRO DE SOLICITANTES PROCESOS] WHERE ID_S = " & aspID_S & " ORDER BY FECHA_P"
            Call CARGAR_DATAGRIDVIEW_2()
            Call ARMAR_DATAGRIDVIEW_2()
        End If
    End Sub
    Private Sub ButtonA_Click(sender As Object, e As EventArgs) Handles ButtonA.Click
        Call VERIFICAR_ACCESOS("072") : If HAY_ACCESO = False Then : Exit Sub : End If
        If aspID_S = 0 Then
            MsgBox("Debe seleccionar un registro Principal de Aspirtante para poder agregar Procesos", vbInformation, "Mensaje del Sistema")
            DGV1.Focus()
            Exit Sub
        End If
        If aspID_ESTADO_R = 2 Then
            MsgBox("El registro Principal del Aspirtante se encuentra Pasivado, no se le pueden agregar Procesos", vbInformation, "Mensaje del Sistema")
            DGV1.Focus()
            Exit Sub
        End If
        Frm_Aspirantes_Add_Procesos.ShowDialog()
        If CERRAR = False Then
            'RemoveHandler Me.DGV2.SelectionChanged, AddressOf DGV2_SelectionChanged
            'Call CARGAR_DATAGRIDVIEW_2()
            'Call ARMAR_DATAGRIDVIEW_2()
            'Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_2()
            'AddHandler Me.DGV2.SelectionChanged, AddressOf DGV2_SelectionChanged
            'RemoveHandler Me.DGV2.SelectionChanged, AddressOf DGV2_SelectionChanged
            'Call CARGAR_DATAGRIDVIEW_2()
            'Call ARMAR_DATAGRIDVIEW_2()
            'Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_2()
            'AddHandler Me.DGV2.SelectionChanged, AddressOf DGV2_SelectionChanged
            'If CERRAR = False Then
            RemoveHandler Me.DGV2.SelectionChanged, AddressOf DGV2_SelectionChanged
            RemoveHandler Me.DGV1.SelectionChanged, AddressOf DGV1_SelectionChanged
            CADENAsql = "SELECT * FROM [VISTA SOL_MAESTRO DE SOLICITANTES] ORDER BY AN"
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_1()
            CADENAsql = "SELECT * FROM [VISTA SOL_MAESTRO DE SOLICITANTES PROCESOS] WHERE ID_S = " & aspID_S & " ORDER BY FECHA_P"
            Call CARGAR_DATAGRIDVIEW_2()
            Call ARMAR_DATAGRIDVIEW_2()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_2()
            AddHandler Me.DGV1.SelectionChanged, AddressOf DGV1_SelectionChanged
            AddHandler Me.DGV2.SelectionChanged, AddressOf DGV2_SelectionChanged
            'End
        End If
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_1()
        Dim I As Integer
        For I = 0 To Me.DGV1.RowCount - 1
            If Me.DGV1.Item(0, I).Value = aspID_S Then
                Me.DGV1.Rows(I).Selected = True
                Me.DGV1.CurrentCell = Me.DGV1.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_2()
        Dim I As Integer
        For I = 0 To Me.DGV2.RowCount - 1
            If Me.DGV2.Item(0, I).Value = asppID_S_P Then
                Me.DGV2.Rows(I).Selected = True
                Me.DGV2.CurrentCell = Me.DGV2.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub DGV2_Click(sender As Object, e As EventArgs) Handles DGV2.Click
        If DGV2.SelectedCells.Count <> 0 Then
            Dim FILA = DGV2.CurrentRow.Index
            If Not IsDBNull(Me.DGV2.Item(0, FILA).Value) Then : asppID_S_P = Me.DGV2.Item(0, FILA).Value : Else : asppID_S_P = 0 : End If
            If Not IsDBNull(Me.DGV2.Item(1, FILA).Value) Then : asppID_S = Me.DGV2.Item(1, FILA).Value : Else : asppID_S = 0 : End If
            If Not IsDBNull(Me.DGV2.Item(2, FILA).Value) Then : asppFECHA_P = Me.DGV2.Item(2, FILA).Value : Else : asppFECHA_P = "" : End If
            If Not IsDBNull(Me.DGV2.Item(3, FILA).Value) Then : asppID_PROC = Me.DGV2.Item(3, FILA).Value : Else : asppID_PROC = 0 : End If
            If Not IsDBNull(Me.DGV2.Item(4, FILA).Value) Then : asppD_PROC = Me.DGV2.Item(4, FILA).Value : Else : asppD_PROC = "" : End If
            If Not IsDBNull(Me.DGV2.Item(5, FILA).Value) Then : asppID_U_RECURSO = Me.DGV2.Item(5, FILA).Value : Else : asppID_U_RECURSO = 0 : End If
            If Not IsDBNull(Me.DGV2.Item(6, FILA).Value) Then : asppD_U_RECURSO = Me.DGV2.Item(6, FILA).Value : Else : asppD_U_RECURSO = "" : End If
            If Not IsDBNull(Me.DGV2.Item(7, FILA).Value) Then : asppREALIZADA = Me.DGV2.Item(7, FILA).Value : Else : asppREALIZADA = False : End If
            If Not IsDBNull(Me.DGV2.Item(8, FILA).Value) Then : asppID_R_SOL = Me.DGV2.Item(8, FILA).Value : Else : asppID_R_SOL = 0 : End If
            If Not IsDBNull(Me.DGV2.Item(9, FILA).Value) Then : asppD_R_SOL = Me.DGV2.Item(9, FILA).Value : Else : asppD_R_SOL = "" : End If
            If Not IsDBNull(Me.DGV2.Item(10, FILA).Value) Then : asppOBSERVACIONES = Me.DGV2.Item(10, FILA).Value : Else : asppOBSERVACIONES = "" : End If
        End If
    End Sub
    Private Sub DGV2_SelectionChanged(sender As Object, e As EventArgs) Handles DGV2.SelectionChanged
        If DGV2.SelectedCells.Count <> 0 Then
            Dim FILA = DGV2.CurrentRow.Index
            If Not IsDBNull(Me.DGV2.Item(0, FILA).Value) Then : asppID_S_P = Me.DGV2.Item(0, FILA).Value : Else : asppID_S_P = 0 : End If
            If Not IsDBNull(Me.DGV2.Item(1, FILA).Value) Then : asppID_S = Me.DGV2.Item(1, FILA).Value : Else : asppID_S = 0 : End If
            If Not IsDBNull(Me.DGV2.Item(2, FILA).Value) Then : asppFECHA_P = Me.DGV2.Item(2, FILA).Value : Else : asppFECHA_P = "" : End If
            If Not IsDBNull(Me.DGV2.Item(3, FILA).Value) Then : asppID_PROC = Me.DGV2.Item(3, FILA).Value : Else : asppID_PROC = 0 : End If
            If Not IsDBNull(Me.DGV2.Item(4, FILA).Value) Then : asppD_PROC = Me.DGV2.Item(4, FILA).Value : Else : asppD_PROC = "" : End If
            If Not IsDBNull(Me.DGV2.Item(5, FILA).Value) Then : asppID_U_RECURSO = Me.DGV2.Item(5, FILA).Value : Else : asppID_U_RECURSO = 0 : End If
            If Not IsDBNull(Me.DGV2.Item(6, FILA).Value) Then : asppD_U_RECURSO = Me.DGV2.Item(6, FILA).Value : Else : asppD_U_RECURSO = "" : End If
            If Not IsDBNull(Me.DGV2.Item(7, FILA).Value) Then : asppREALIZADA = Me.DGV2.Item(7, FILA).Value : Else : asppREALIZADA = False : End If
            If Not IsDBNull(Me.DGV2.Item(8, FILA).Value) Then : asppID_R_SOL = Me.DGV2.Item(8, FILA).Value : Else : asppID_R_SOL = 0 : End If
            If Not IsDBNull(Me.DGV2.Item(9, FILA).Value) Then : asppD_R_SOL = Me.DGV2.Item(9, FILA).Value : Else : asppD_R_SOL = "" : End If
            If Not IsDBNull(Me.DGV2.Item(10, FILA).Value) Then : asppOBSERVACIONES = Me.DGV2.Item(10, FILA).Value : Else : asppOBSERVACIONES = "" : End If
            If Not IsDBNull(Me.DGV2.Item(11, FILA).Value) Then : asppID_ESTADO_R = Me.DGV2.Item(11, FILA).Value : Else : asppID_ESTADO_R = 0 : End If
        End If
    End Sub
    Private Sub ButtonB_Click(sender As Object, e As EventArgs) Handles ButtonB.Click
        Call VERIFICAR_ACCESOS("073") : If HAY_ACCESO = False Then : Exit Sub : End If
        If asppID_ESTADO_R = 2 Then
            MsgBox("El registro ya se encuentra Pasivado", vbInformation, "Mensaje del Sistema")
            Me.DGV2.Focus()
            Exit Sub
        End If
        If aspID_S = 0 Then
            MsgBox("Debe seleccionar un registro Principal de Aspirante para poder agregar Procesos", vbInformation, "Mensaje del Sistema")
            DGV1.Focus()
            Exit Sub
        End If
        If aspID_ESTADO_R = 2 Then
            MsgBox("El registro Principal del Aspirante se encuentra Pasivado, no se le pueden agregar Procesos", vbInformation, "Mensaje del Sistema")
            DGV1.Focus()
            Exit Sub
        End If
        If asppID_S_P = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            DGV2.Focus()
            Exit Sub
        End If
        Frm_Aspirantes_Upd_Procesos.ShowDialog()
        If CERRAR = False Then

            'RemoveHandler Me.DGV2.SelectionChanged, AddressOf DGV2_SelectionChanged
            'Call CARGAR_DATAGRIDVIEW_2()
            'Call ARMAR_DATAGRIDVIEW_2()
            'Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_2()
            'AddHandler Me.DGV2.SelectionChanged, AddressOf DGV2_SelectionChanged
            'If CERRAR = False Then
            RemoveHandler Me.DGV2.SelectionChanged, AddressOf DGV2_SelectionChanged
            RemoveHandler Me.DGV1.SelectionChanged, AddressOf DGV1_SelectionChanged
            CADENAsql = "SELECT * FROM [VISTA SOL_MAESTRO DE SOLICITANTES] ORDER BY AN"
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_1()
            CADENAsql = "SELECT * FROM [VISTA SOL_MAESTRO DE SOLICITANTES PROCESOS] WHERE ID_S = " & aspID_S & " ORDER BY FECHA_P"
            Call CARGAR_DATAGRIDVIEW_2()
            Call ARMAR_DATAGRIDVIEW_2()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_2()
            AddHandler Me.DGV1.SelectionChanged, AddressOf DGV1_SelectionChanged
            AddHandler Me.DGV2.SelectionChanged, AddressOf DGV2_SelectionChanged
            'End
        End If
    End Sub
    Private Sub DGV2_DoubleClick(sender As Object, e As EventArgs) Handles DGV2.DoubleClick
        Call VERIFICAR_ACCESOS("073") : If HAY_ACCESO = False Then : Exit Sub : End If
        If asppID_ESTADO_R = 2 Then
            MsgBox("El registro ya se encuentra Pasivado", vbInformation, "Mensaje del Sistema")
            Me.DGV2.Focus()
            Exit Sub
        End If
        If aspID_S = 0 Then
            MsgBox("Debe seleccionar un registro Principal de Aspirante para poder agregar Procesos", vbInformation, "Mensaje del Sistema")
            DGV1.Focus()
            Exit Sub
        End If
        If aspID_ESTADO_R = 2 Then
            MsgBox("El registro Principal del Aspirante se encuentra Pasivado, no se le pueden agregar Procesos", vbInformation, "Mensaje del Sistema")
            DGV1.Focus()
            Exit Sub
        End If
        If asppID_S_P = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            DGV2.Focus()
            Exit Sub
        End If
        Frm_Aspirantes_Upd_Procesos.ShowDialog()
        If CERRAR = False Then

            'RemoveHandler Me.DGV2.SelectionChanged, AddressOf DGV2_SelectionChanged
            'Call CARGAR_DATAGRIDVIEW_2()
            'Call ARMAR_DATAGRIDVIEW_2()
            'Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_2()
            'AddHandler Me.DGV2.SelectionChanged, AddressOf DGV2_SelectionChanged
            'If CERRAR = False Then
            RemoveHandler Me.DGV2.SelectionChanged, AddressOf DGV2_SelectionChanged
            RemoveHandler Me.DGV1.SelectionChanged, AddressOf DGV1_SelectionChanged
            CADENAsql = "SELECT * FROM [VISTA SOL_MAESTRO DE SOLICITANTES] ORDER BY AN"
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_1()
            CADENAsql = "SELECT * FROM [VISTA SOL_MAESTRO DE SOLICITANTES PROCESOS] WHERE ID_S = " & aspID_S & " ORDER BY FECHA_P"
            Call CARGAR_DATAGRIDVIEW_2()
            Call ARMAR_DATAGRIDVIEW_2()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_2()
            AddHandler Me.DGV1.SelectionChanged, AddressOf DGV1_SelectionChanged
            AddHandler Me.DGV2.SelectionChanged, AddressOf DGV2_SelectionChanged
            'End If
        End If
    End Sub
    Private Sub ButtonC_Click(sender As Object, e As EventArgs) Handles ButtonC.Click
        Call VERIFICAR_ACCESOS("074") : If HAY_ACCESO = False Then : Exit Sub : End If
        If asppID_ESTADO_R = 2 Then
            MsgBox("El registro ya se encuentra Pasivado", vbInformation, "Mensaje del Sistema")
            Me.DGV2.Focus()
            Exit Sub
        End If
        If aspID_S = 0 Then
            MsgBox("Debe seleccionar un registro Principal de Aspirante para poder agregar Procesos", vbInformation, "Mensaje del Sistema")
            DGV1.Focus()
            Exit Sub
        End If
        If aspID_ESTADO_R = 2 Then
            MsgBox("El registro Principal del Aspirante se encuentra Pasivado, no puede Actualizarse el registro actual", vbInformation, "Mensaje del Sistema")
            DGV1.Focus()
            Exit Sub
        End If
        If asppID_S_P = 0 Then
            MsgBox("Debe seleccionar el registro que desea Pasivar", vbInformation, "Mensaje del Sistema")
            DGV2.Focus()
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Esta seguro de Pasivar el registro seleccionado?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            Call PASIVAR_REGISTRO_2()
            RemoveHandler Me.DGV2.SelectionChanged, AddressOf DGV2_SelectionChanged
            Call CARGAR_DATAGRIDVIEW_2()
            Call ARMAR_DATAGRIDVIEW_2()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_2()
            AddHandler Me.DGV2.SelectionChanged, AddressOf DGV2_SelectionChanged
        Else
            Exit Sub
        End If
    End Sub
    Private Sub PASIVAR_REGISTRO_2()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [SOL_MAESTRO DE SOLICITANTES PROCESOS] SET ID_ESTADO_R = 2 WHERE ID_S_P = " & CInt(asppID_S_P) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub DGV2_RowPrePaint(sender As Object, e As DataGridViewRowPrePaintEventArgs) Handles DGV2.RowPrePaint
        Call FORMATO_FILAS_DATAGRIDVIEW_2()
    End Sub
    Private Sub FORMATO_FILAS_DATAGRIDVIEW_2()
        Dim style As New DataGridViewCellStyle
        style.Font = New Font(DGV2.Font, FontStyle.Bold)
        For Each Row As DataGridViewRow In DGV2.Rows
            If Row.Cells(11).Value = "2" Then
                Row.DefaultCellStyle.BackColor = Color.DimGray
            End If
        Next
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call VERIFICAR_ACCESOS("065") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Aspirantes_Buscar.ShowDialog()
        If CERRAR = False Then
            RemoveHandler Me.DGV2.SelectionChanged, AddressOf DGV2_SelectionChanged
            RemoveHandler Me.DGV1.SelectionChanged, AddressOf DGV1_SelectionChanged
            CADENAsql = "SELECT * FROM [VISTA SOL_MAESTRO DE SOLICITANTES] ORDER BY AN"
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_1()
            CADENAsql = "SELECT * FROM [VISTA SOL_MAESTRO DE SOLICITANTES PROCESOS] WHERE ID_S = " & aspID_S & " ORDER BY FECHA_P"
            Call CARGAR_DATAGRIDVIEW_2()
            Call ARMAR_DATAGRIDVIEW_2()
            AddHandler Me.DGV1.SelectionChanged, AddressOf DGV1_SelectionChanged
            AddHandler Me.DGV2.SelectionChanged, AddressOf DGV2_SelectionChanged
        End If
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Call VERIFICAR_ACCESOS("068") : If HAY_ACCESO = False Then : Exit Sub : End If
        If aspID_S = 0 Then
            MsgBox("Debe seleccionar un registro Principal de Aspirante para poder Pasivarlo", vbInformation, "Mensaje del Sistema")
            DGV1.Focus()
            Exit Sub
        End If
        If aspID_ESTADO_R = 2 Then
            MsgBox("El registro ya se encuentra Pasivado", vbInformation, "Mensaje del Sistema")
            DGV1.Focus()
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Esta seguro de Pasivar el registro seleccionado?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            Call PASIVAR_REGISTRO_1()
            RemoveHandler Me.DGV2.SelectionChanged, AddressOf DGV2_SelectionChanged
            RemoveHandler Me.DGV1.SelectionChanged, AddressOf DGV1_SelectionChanged
            CADENAsql = "SELECT * FROM [VISTA SOL_MAESTRO DE SOLICITANTES] ORDER BY AN"
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_1()
            CADENAsql = "SELECT * FROM [VISTA SOL_MAESTRO DE SOLICITANTES PROCESOS] WHERE ID_S = " & aspID_S & " ORDER BY FECHA_P"
            Call CARGAR_DATAGRIDVIEW_2()
            Call ARMAR_DATAGRIDVIEW_2()
            AddHandler Me.DGV1.SelectionChanged, AddressOf DGV1_SelectionChanged
            AddHandler Me.DGV2.SelectionChanged, AddressOf DGV2_SelectionChanged
        Else
            Exit Sub
        End If
    End Sub
    Private Sub PASIVAR_REGISTRO_1()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [SOL_MAESTRO DE SOLICITANTES] SET ID_ESTADO_R = 2 WHERE ID_S = " & CInt(aspID_S) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub DGV1_RowPrePaint(sender As Object, e As DataGridViewRowPrePaintEventArgs) Handles DGV1.RowPrePaint
        Call FORMATO_FILAS_DATAGRIDVIEW_1()
    End Sub
    Private Sub FORMATO_FILAS_DATAGRIDVIEW_1()
        Dim style As New DataGridViewCellStyle
        style.Font = New Font(DGV1.Font, FontStyle.Bold)
        For Each Row As DataGridViewRow In DGV1.Rows
            If Row.Cells(23).Value = "2" Then
                Row.DefaultCellStyle.BackColor = Color.DimGray
            End If
        Next
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call VERIFICAR_ACCESOS("066") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Aspirantes_Add.ShowDialog()
        If CERRAR = False Then
            RemoveHandler Me.DGV2.SelectionChanged, AddressOf DGV2_SelectionChanged
            RemoveHandler Me.DGV1.SelectionChanged, AddressOf DGV1_SelectionChanged
            CADENAsql = "SELECT * FROM [VISTA SOL_MAESTRO DE SOLICITANTES] ORDER BY AN"
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_1()
            CADENAsql = "SELECT * FROM [VISTA SOL_MAESTRO DE SOLICITANTES PROCESOS] WHERE ID_S = " & aspID_S & " ORDER BY FECHA_P"
            Call CARGAR_DATAGRIDVIEW_2()
            Call ARMAR_DATAGRIDVIEW_2()
            AddHandler Me.DGV1.SelectionChanged, AddressOf DGV1_SelectionChanged
            AddHandler Me.DGV2.SelectionChanged, AddressOf DGV2_SelectionChanged
            If AUTOMATICO = True Then
                Frm_Aspirantes_Add_Procesos.ShowDialog()
                If CERRAR = False Then
                    RemoveHandler Me.DGV2.SelectionChanged, AddressOf DGV2_SelectionChanged
                    Call CARGAR_DATAGRIDVIEW_2()
                    Call ARMAR_DATAGRIDVIEW_2()
                    Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_2()
                    AddHandler Me.DGV2.SelectionChanged, AddressOf DGV2_SelectionChanged
                End If
            End If
        Else
            Exit Sub
        End If
    End Sub
    Private Sub DGV1_DoubleClick(sender As Object, e As EventArgs) Handles DGV1.DoubleClick
        Call VERIFICAR_ACCESOS("067") : If HAY_ACCESO = False Then : Exit Sub : End If
        If aspID_S = 0 Then
            MsgBox("Debe seleccionar un registro Principal de Aspirante para poder Editarlo", vbInformation, "Mensaje del Sistema")
            DGV1.Focus()
            Exit Sub
        End If
        If aspID_ESTADO_R = 2 Then
            MsgBox("El registro se encuentra Pasivado", vbInformation, "Mensaje del Sistema")
            DGV1.Focus()
            Exit Sub
        End If
        Frm_Aspirantes_Upd.ShowDialog()
        If CERRAR = False Then
            RemoveHandler Me.DGV2.SelectionChanged, AddressOf DGV2_SelectionChanged
            RemoveHandler Me.DGV1.SelectionChanged, AddressOf DGV1_SelectionChanged
            CADENAsql = "SELECT * FROM [VISTA SOL_MAESTRO DE SOLICITANTES] ORDER BY AN"
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_1()
            CADENAsql = "SELECT * FROM [VISTA SOL_MAESTRO DE SOLICITANTES PROCESOS] WHERE ID_S = " & aspID_S & " ORDER BY FECHA_P"
            Call CARGAR_DATAGRIDVIEW_2()
            Call ARMAR_DATAGRIDVIEW_2()
            AddHandler Me.DGV1.SelectionChanged, AddressOf DGV1_SelectionChanged
            AddHandler Me.DGV2.SelectionChanged, AddressOf DGV2_SelectionChanged
        Else
            Exit Sub
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call VERIFICAR_ACCESOS("067") : If HAY_ACCESO = False Then : Exit Sub : End If
        If aspID_S = 0 Then
            MsgBox("Debe seleccionar un registro Principal de Aspirante para poder Editarlo", vbInformation, "Mensaje del Sistema")
            DGV1.Focus()
            Exit Sub
        End If
        If aspID_ESTADO_R = 2 Then
            MsgBox("El registro se encuentra Pasivado", vbInformation, "Mensaje del Sistema")
            DGV1.Focus()
            Exit Sub
        End If
        Frm_Aspirantes_Upd.ShowDialog()
        If CERRAR = False Then
            RemoveHandler Me.DGV2.SelectionChanged, AddressOf DGV2_SelectionChanged
            RemoveHandler Me.DGV1.SelectionChanged, AddressOf DGV1_SelectionChanged
            CADENAsql = "SELECT * FROM [VISTA SOL_MAESTRO DE SOLICITANTES] ORDER BY AN"
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_1()
            CADENAsql = "SELECT * FROM [VISTA SOL_MAESTRO DE SOLICITANTES PROCESOS] WHERE ID_S = " & aspID_S & " ORDER BY FECHA_P"
            Call CARGAR_DATAGRIDVIEW_2()
            Call ARMAR_DATAGRIDVIEW_2()
            AddHandler Me.DGV1.SelectionChanged, AddressOf DGV1_SelectionChanged
            AddHandler Me.DGV2.SelectionChanged, AddressOf DGV2_SelectionChanged
        Else
            Exit Sub
        End If
    End Sub
    Private Sub Button51_Click(sender As Object, e As EventArgs) Handles Button51.Click
        Call VERIFICAR_ACCESOS("070") : If HAY_ACCESO = False Then : Exit Sub : End If
        If aspID_S = 0 Then
            MsgBox("Debe seleccionar el registro del Aspirante que desea Imprimir", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Solicitudes_de_Aspirantes_PDF.ShowDialog()
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Call VERIFICAR_ACCESOS("069") : If HAY_ACCESO = False Then : Exit Sub : End If
        If aspID_S = 0 Then
            MsgBox("Debe seleccionar el registro del Aspirante que desea Imprimir", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Aspirantes_Imprimir.ShowDialog()
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Call VERIFICAR_ACCESOS("071") : If HAY_ACCESO = False Then : Exit Sub : End If
        If aspID_S = 0 Or aspAN = "" Then
            MsgBox("Debe seleccionar el registro del Aspirante que desea Imprimir", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        If aspID_ESTADO_R = 2 Then
            MsgBox("El registro se encuentra Pasivado", vbInformation, "Mensaje del Sistema")
            DGV1.Focus()
            Exit Sub
        End If
        Frm_Documentos.ShowDialog()
    End Sub
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        'If Len(Me.TextBox2.Text) = 0 Then
        '    MsgBox("Debe digitar correctamente el Texto que desea Buscar", vbInformation, "Mensaje del sistema")
        '    Me.TextBox2.Focus()
        'Else
        RemoveHandler Me.DGV1.SelectionChanged, AddressOf DGV1_SelectionChanged
        RemoveHandler Me.DGV2.RowPrePaint, AddressOf DGV2_RowPrePaint
        Call CARGAR_DATAGRIDVIEW_FILTRO()
        Call ARMAR_DATAGRIDVIEW()
        CADENAsql = "SELECT * FROM [VISTA SOL_MAESTRO DE SOLICITANTES PROCESOS] WHERE ID_S = " & aspID_S & " ORDER BY FECHA_P"
        Call CARGAR_DATAGRIDVIEW_2()
        Call ARMAR_DATAGRIDVIEW_2()
        AddHandler Me.DGV2.RowPrePaint, AddressOf DGV2_RowPrePaint
        AddHandler Me.DGV1.SelectionChanged, AddressOf DGV1_SelectionChanged
        'End If
    End Sub
    'Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
    '    If Len(Me.TextBox2.Text) = 0 Then
    '        MsgBox("Debe digitar correctamente el Texto que desea Buscar", vbInformation, "Mensaje del sistema")
    '        Me.TextBox2.Focus()
    '    Else
    '        RemoveHandler Me.DGV1.SelectionChanged, AddressOf DGV1_SelectionChanged
    '        RemoveHandler Me.DGV2.RowPrePaint, AddressOf DGV2_RowPrePaint
    '        Call CARGAR_DATAGRIDVIEW_FILTRO()
    '        Call ARMAR_DATAGRIDVIEW()
    '        CADENAsql = "SELECT * FROM [VISTA SOL_MAESTRO DE SOLICITANTES PROCESOS] WHERE ID_S = " & aspID_S & " ORDER BY FECHA_P"
    '        Call CARGAR_DATAGRIDVIEW_2()
    '        Call ARMAR_DATAGRIDVIEW_2()
    '        AddHandler Me.DGV2.RowPrePaint, AddressOf DGV2_RowPrePaint
    '        AddHandler Me.DGV1.SelectionChanged, AddressOf DGV1_SelectionChanged
    '    End If
    'End Sub
    Private Sub CARGAR_DATAGRIDVIEW_FILTRO_2()
        CADENAsql = "SELECT * FROM [VISTA SOL_MAESTRO DE SOLICITANTES] WHERE (FECHA_I BETWEEN " & FECHA1 & " AND " & FECHA2 & ") ORDER BY FECHA_I"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA SOL_MAESTRO DE SOLICITANTES]")
        DGV1.DataSource = MiDataSet.Tables("[VISTA SOL_MAESTRO DE SOLICITANTES]")
        Me.Label2.Text = DGV1.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_FILTRO()
        CADENAsql = "Select * from [VISTA SOL_MAESTRO DE SOLICITANTES] WHERE (PERFIL LIKE '%" & Me.TextBox2.Text & "%') ORDER BY PERFIL"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA SOL_MAESTRO DE SOLICITANTES]")
        DGV1.DataSource = MiDataSet.Tables("[VISTA SOL_MAESTRO DE SOLICITANTES]")
        Me.Label2.Text = DGV1.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub CheckBox1_CheckedChanged_1(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If Me.CheckBox1.Checked = True Then
            'Me.TextBox2.Enabled = True
            'Me.DateTimePicker1.Enabled = False
            'Me.DateTimePicker2.Enabled = False
        Else
            Me.TextBox2.Enabled = False
            Me.TextBox2.Text = ""
            RemoveHandler Me.DGV1.SelectionChanged, AddressOf DGV1_SelectionChanged
            RemoveHandler Me.DGV2.RowPrePaint, AddressOf DGV2_RowPrePaint
            CADENAsql = "SELECT * FROM [VISTA SOL_MAESTRO DE SOLICITANTES] ORDER BY AN"
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            CADENAsql = "SELECT * FROM [VISTA SOL_MAESTRO DE SOLICITANTES PROCESOS] WHERE ID_S = " & aspID_S & " ORDER BY FECHA_P"
            Call CARGAR_DATAGRIDVIEW_2()
            Call ARMAR_DATAGRIDVIEW_2()
            AddHandler Me.DGV2.RowPrePaint, AddressOf DGV2_RowPrePaint
            AddHandler Me.DGV1.SelectionChanged, AddressOf DGV1_SelectionChanged
        End If
    End Sub
    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If Me.CheckBox2.Checked = True Then
            'Me.TextBox2.Enabled = False
            'Me.DateTimePicker1.Enabled = True
            'Me.DateTimePicker2.Enabled = True
        Else
            Me.DateTimePicker1.Enabled = False
            Me.DateTimePicker2.Enabled = False
            RemoveHandler Me.DGV1.SelectionChanged, AddressOf DGV1_SelectionChanged
            RemoveHandler Me.DGV2.RowPrePaint, AddressOf DGV2_RowPrePaint
            CADENAsql = "SELECT * FROM [VISTA SOL_MAESTRO DE SOLICITANTES] ORDER BY AN"
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            CADENAsql = "SELECT * FROM [VISTA SOL_MAESTRO DE SOLICITANTES PROCESOS] WHERE ID_S = " & aspID_S & " ORDER BY FECHA_P"
            Call CARGAR_DATAGRIDVIEW_2()
            Call ARMAR_DATAGRIDVIEW_2()
            AddHandler Me.DGV2.RowPrePaint, AddressOf DGV2_RowPrePaint
            AddHandler Me.DGV1.SelectionChanged, AddressOf DGV1_SelectionChanged
        End If
    End Sub
    Private Sub CheckBox1_Click(sender As Object, e As EventArgs) Handles CheckBox1.Click
        If Me.CheckBox1.Checked = True Then
            Me.TextBox2.Enabled = True
            Me.DateTimePicker1.Enabled = False
            Me.DateTimePicker2.Enabled = False
            Me.Button10.Enabled = False
            Me.CheckBox2.Checked = False
        Else
            Me.TextBox2.Enabled = False
            Me.DateTimePicker1.Enabled = True
            Me.DateTimePicker2.Enabled = True
            Me.Button10.Enabled = True
        End If
    End Sub
    Private Sub CheckBox2_Click(sender As Object, e As EventArgs) Handles CheckBox2.Click
        If Me.CheckBox2.Checked = True Then
            Me.TextBox2.Enabled = False
            Me.DateTimePicker1.Enabled = True
            Me.DateTimePicker2.Enabled = True
            Me.Button10.Enabled = True
            Me.CheckBox1.Checked = False
        Else
            Me.TextBox2.Enabled = True
            Me.DateTimePicker1.Enabled = False
            Me.DateTimePicker2.Enabled = False
            Me.Button10.Enabled = False
        End If
    End Sub
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        'Me.DateTimePicker1.Value = Mid(vFECHA_A, 1, 10)
        'Me.DateTimePicker2.Value = Mid(vFECHA_A, 1, 10)
        FECHA01 = Me.DateTimePicker1.Text
        FECHA1 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA01 + "', 105), 23)"
        FECHA02 = Me.DateTimePicker2.Text
        FECHA2 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA02 + "', 105), 23)"
        If Me.CheckBox2.Checked = True Then
            RemoveHandler Me.DGV1.SelectionChanged, AddressOf DGV1_SelectionChanged
            RemoveHandler Me.DGV2.RowPrePaint, AddressOf DGV2_RowPrePaint
            Call CARGAR_DATAGRIDVIEW_FILTRO_2()
            Call ARMAR_DATAGRIDVIEW()
            CADENAsql = "SELECT * FROM [VISTA SOL_MAESTRO DE SOLICITANTES PROCESOS] WHERE ID_S = " & aspID_S & " ORDER BY FECHA_P"
            Call CARGAR_DATAGRIDVIEW_2()
            Call ARMAR_DATAGRIDVIEW_2()
            AddHandler Me.DGV2.RowPrePaint, AddressOf DGV2_RowPrePaint
            AddHandler Me.DGV1.SelectionChanged, AddressOf DGV1_SelectionChanged
        End If
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Try
            TITULO_EXCEL = "Aspirantes (Reclutamiento, Selección y Contratación)"
            EXPORTAR_DATOS_A_EXCEL(DGV1, "(" & UCase(Now) & ")")
        Catch ex As Exception
            MessageBox.Show("No se puede generar el documento Excel.", "Mensaje del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    'Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
    '    Me.TextBox2.Text = ""
    '    RemoveHandler Me.DGV1.SelectionChanged, AddressOf DGV1_SelectionChanged
    '    RemoveHandler Me.DGV2.RowPrePaint, AddressOf DGV2_RowPrePaint
    '    CADENAsql = "SELECT * FROM [VISTA SOL_MAESTRO DE SOLICITANTES] ORDER BY AN"
    '    Call CARGAR_DATAGRIDVIEW()
    '    Call ARMAR_DATAGRIDVIEW()
    '    CADENAsql = "SELECT * FROM [VISTA SOL_MAESTRO DE SOLICITANTES PROCESOS] WHERE ID_S = " & aspID_S & " ORDER BY FECHA_P"
    '    Call CARGAR_DATAGRIDVIEW_2()
    '    Call ARMAR_DATAGRIDVIEW_2()
    '    AddHandler Me.DGV2.RowPrePaint, AddressOf DGV2_RowPrePaint
    '    AddHandler Me.DGV1.SelectionChanged, AddressOf DGV1_SelectionChanged
    'End Sub
End Class