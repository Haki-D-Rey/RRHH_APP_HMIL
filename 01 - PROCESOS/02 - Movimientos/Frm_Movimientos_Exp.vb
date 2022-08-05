Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class Frm_Movimientos_Exp
    Private Sub Frm_Movimientos_Exp_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            MOVIMIENTOS = False
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Movimientos_Exp_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Frm_Principal.Timer1.Stop()
        MOVIMIENTOS = True
        Me.Cursor = Cursors.WaitCursor
        RemoveHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
        RemoveHandler Me.DateTimePicker1.ValueChanged, AddressOf DateTimePicker1_ValueChanged
        RemoveHandler Me.DateTimePicker2.ValueChanged, AddressOf DateTimePicker2_ValueChanged
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_MOVIMIENTO_DE_EXP()
        Me.DateTimePicker1.Value = Now : FECHA01 = Me.DateTimePicker1.Text
        FECHA1 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA01 + "', 105), 23)"
        Me.DateTimePicker2.Value = Now : FECHA02 = Me.DateTimePicker2.Text
        FECHA2 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA02 + "', 105), 23)"

        If isuID_TIPO_USUARIO = 5 Then
            CADENAsql = "SELECT * FROM [dbo].[VISTA HISTORICO DE MOVIMIENTO] WHERE FEC_MOV BETWEEN " & FECHA1 & " AND " & FECHA2 & " AND ID_T_M_EXP = " & Me.ComboBox1.SelectedValue & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY FEC_MOV, ID_M_P"
        Else
            CADENAsql = "SELECT * FROM [dbo].[VISTA HISTORICO DE MOVIMIENTO] WHERE FEC_MOV BETWEEN " & FECHA1 & " AND " & FECHA2 & " AND ID_T_M_EXP = " & Me.ComboBox1.SelectedValue & " ORDER BY FEC_MOV, ID_M_P"
        End If

        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
        AddHandler Me.DateTimePicker1.ValueChanged, AddressOf DateTimePicker1_ValueChanged
        AddHandler Me.DateTimePicker2.ValueChanged, AddressOf DateTimePicker2_ValueChanged
        AddHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        MOVIMIENTOS = False
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_MOVIMIENTO_DE_EXP()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE MOVIMIENTO DE EXP] ORDER BY ID_T_M_EXP", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_T_M_EXP"
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
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA HISTORICO DE MOVIMIENTO]")
        DGV.DataSource = MiDataSet.Tables("[VISTA HISTORICO DE MOVIMIENTO]")
        Me.Label4.Text = DGV.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW()
        Me.DGV.Columns(0).HeaderText = "Id"
        Me.DGV.Columns(0).Width = 10
        Me.DGV.Columns(0).Visible = True
        Me.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(1).HeaderText = "Fecha Ing. PAME"
        Me.DGV.Columns(1).Width = 65
        Me.DGV.Columns(1).Visible = True
        Me.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(2).HeaderText = "ID_M_P"
        Me.DGV.Columns(2).Width = 0
        Me.DGV.Columns(2).Visible = False

        Me.DGV.Columns(3).HeaderText = "Codigo"
        Me.DGV.Columns(3).Width = 70
        Me.DGV.Columns(3).Visible = True
        Me.DGV.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(4).HeaderText = "Apellidos y Nombres"
        Me.DGV.Columns(4).Width = 145
        Me.DGV.Columns(4).Visible = True
        Me.DGV.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(5).HeaderText = "Fecha Grabacion"
        Me.DGV.Columns(5).Width = 65
        Me.DGV.Columns(5).Visible = True
        Me.DGV.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(6).HeaderText = "Fecha Aplicacion"
        Me.DGV.Columns(6).Width = 65
        Me.DGV.Columns(6).Visible = True
        Me.DGV.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(7).HeaderText = "ID_T_M_EXP"
        Me.DGV.Columns(7).Width = 0
        Me.DGV.Columns(7).Visible = False

        Me.DGV.Columns(8).HeaderText = "Tipo Movimiento"
        Me.DGV.Columns(8).Width = 0
        Me.DGV.Columns(8).Visible = False

        Me.DGV.Columns(9).HeaderText = "ID_ST_M_L_EXP"
        Me.DGV.Columns(9).Width = 0
        Me.DGV.Columns(9).Visible = False

        Me.DGV.Columns(10).HeaderText = "Causa Legal"
        Me.DGV.Columns(10).Width = 55
        Me.DGV.Columns(10).Visible = True
        Me.DGV.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(11).HeaderText = "ID_ST_M_R_EXP"
        Me.DGV.Columns(11).Width = 0
        Me.DGV.Columns(11).Visible = False

        Me.DGV.Columns(12).HeaderText = "Causa Real"
        Me.DGV.Columns(12).Width = 55
        Me.DGV.Columns(12).Visible = True
        Me.DGV.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(13).HeaderText = "Cargo"
        Me.DGV.Columns(13).Width = 70
        Me.DGV.Columns(13).Visible = True
        Me.DGV.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(14).HeaderText = "Gp"
        Me.DGV.Columns(14).Width = 0
        Me.DGV.Columns(14).Visible = False
        Me.DGV.Columns(14).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(15).HeaderText = "Gr"
        Me.DGV.Columns(15).Width = 30
        Me.DGV.Columns(15).Visible = True
        Me.DGV.Columns(15).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(16).HeaderText = "Nivel 1"
        Me.DGV.Columns(16).Width = 60
        Me.DGV.Columns(16).Visible = True
        Me.DGV.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(17).HeaderText = "Nivel 2"
        Me.DGV.Columns(17).Width = 60
        Me.DGV.Columns(17).Visible = True
        Me.DGV.Columns(17).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(18).HeaderText = "Nivel 3"
        Me.DGV.Columns(18).Width = 60
        Me.DGV.Columns(18).Visible = True
        Me.DGV.Columns(18).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(19).HeaderText = "Nivel 4"
        Me.DGV.Columns(19).Width = 60
        Me.DGV.Columns(19).Visible = True
        Me.DGV.Columns(19).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(20).HeaderText = "Nivel 5"
        Me.DGV.Columns(20).Width = 60
        Me.DGV.Columns(20).Visible = True
        Me.DGV.Columns(20).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(21).HeaderText = "Nivel 6"
        Me.DGV.Columns(21).Width = 60
        Me.DGV.Columns(21).Visible = True
        Me.DGV.Columns(21).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(22).HeaderText = "Nivel 7"
        Me.DGV.Columns(22).Width = 60
        Me.DGV.Columns(22).Visible = True
        Me.DGV.Columns(22).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(23).HeaderText = "Nivel 8"
        Me.DGV.Columns(23).Width = 60
        Me.DGV.Columns(23).Visible = True
        Me.DGV.Columns(23).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(24).HeaderText = "N_ORDEN"
        Me.DGV.Columns(24).Width = 0
        Me.DGV.Columns(24).Visible = False

        Me.DGV.Columns(25).HeaderText = "Nomina"
        Me.DGV.Columns(25).Width = 0
        Me.DGV.Columns(25).Visible = False
        Me.DGV.Columns(25).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(26).HeaderText = "Salario"
        Me.DGV.Columns(26).Width = 0
        Me.DGV.Columns(26).Visible = False
        Me.DGV.Columns(26).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(26).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV.Columns(26).DefaultCellStyle.Format = "##,##0.00"

        Me.DGV.Columns(27).HeaderText = "Detalle"
        Me.DGV.Columns(27).Width = 0
        Me.DGV.Columns(27).Visible = False
        Me.DGV.Columns(27).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(28).HeaderText = "Mixta"
        Me.DGV.Columns(28).Width = 0
        Me.DGV.Columns(28).Visible = False
        Me.DGV.Columns(28).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(29).HeaderText = "Militar"
        Me.DGV.Columns(29).Width = 0
        Me.DGV.Columns(29).Visible = False
        Me.DGV.Columns(29).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(30).HeaderText = "PAME"
        Me.DGV.Columns(30).Width = 0
        Me.DGV.Columns(30).Visible = False
        Me.DGV.Columns(30).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(31).HeaderText = "Apli cado"
        Me.DGV.Columns(31).Width = 0
        Me.DGV.Columns(31).Visible = False
        Me.DGV.Columns(31).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(32).HeaderText = "Estadis tico"
        Me.DGV.Columns(32).Width = 0
        Me.DGV.Columns(32).Visible = False
        Me.DGV.Columns(32).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(33).HeaderText = "USUARIO_ACT"
        Me.DGV.Columns(33).Width = 0
        Me.DGV.Columns(33).Visible = False

        Me.DGV.Columns(34).HeaderText = "FECHA_ACT"
        Me.DGV.Columns(34).Width = 0
        Me.DGV.Columns(34).Visible = False

        Me.DGV.Columns(35).HeaderText = "D_ADMIN"
        Me.DGV.Columns(35).Width = 0
        Me.DGV.Columns(35).Visible = False
    End Sub
    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        FECHA01 = Me.DateTimePicker1.Text
        Dim FECHA1 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA01 + "', 105), 23)"
        If isuID_TIPO_USUARIO = 5 Then
            CADENAsql = "SELECT * FROM [dbo].[VISTA HISTORICO DE MOVIMIENTO] WHERE FEC_MOV BETWEEN " & FECHA1 & " AND " & FECHA2 & " AND ID_T_M_EXP = " & Me.ComboBox1.SelectedValue & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY FEC_MOV, ID_M_P"
        Else
            CADENAsql = "SELECT * FROM [dbo].[VISTA HISTORICO DE MOVIMIENTO] WHERE FEC_MOV BETWEEN " & FECHA1 & " AND " & FECHA2 & " AND ID_T_M_EXP = " & Me.ComboBox1.SelectedValue & " ORDER BY FEC_MOV, ID_M_P"
        End If
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
    End Sub
    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        FECHA02 = Me.DateTimePicker2.Text
        FECHA2 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA02 + "', 105), 23)"
        If isuID_TIPO_USUARIO = 5 Then
            CADENAsql = "SELECT * FROM [dbo].[VISTA HISTORICO DE MOVIMIENTO] WHERE FEC_MOV BETWEEN " & FECHA1 & " AND " & FECHA2 & " AND ID_T_M_EXP = " & Me.ComboBox1.SelectedValue & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY FEC_MOV, ID_M_P"
        Else
            CADENAsql = "SELECT * FROM [dbo].[VISTA HISTORICO DE MOVIMIENTO] WHERE FEC_MOV BETWEEN " & FECHA1 & " AND " & FECHA2 & " AND ID_T_M_EXP = " & Me.ComboBox1.SelectedValue & " ORDER BY FEC_MOV, ID_M_P"
        End If
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        Dim I As Integer
        For I = 0 To Me.DGV.RowCount - 1
            If Me.DGV.Item(0, I).Value = vID_HIST_M Then
                Me.DGV.Rows(I).Selected = True
                Me.DGV.CurrentCell = Me.DGV.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        FECHA01 = Me.DateTimePicker1.Text
        FECHA1 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA01 + "', 105), 23)"
        FECHA02 = Me.DateTimePicker2.Text
        FECHA2 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA02 + "', 105), 23)"
        If isuID_TIPO_USUARIO = 5 Then
            CADENAsql = "SELECT * FROM [dbo].[VISTA HISTORICO DE MOVIMIENTO] WHERE FEC_MOV BETWEEN " & FECHA1 & " AND " & FECHA2 & " AND ID_T_M_EXP = " & Me.ComboBox1.SelectedValue & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY FEC_MOV, ID_M_P"
        Else
            CADENAsql = "SELECT * FROM [dbo].[VISTA HISTORICO DE MOVIMIENTO] WHERE FEC_MOV BETWEEN " & FECHA1 & " AND " & FECHA2 & " AND ID_T_M_EXP = " & Me.ComboBox1.SelectedValue & " ORDER BY FEC_MOV, ID_M_P"
        End If
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
    End Sub
    Private Sub DGV_RowPrePaint(sender As Object, e As DataGridViewRowPrePaintEventArgs) Handles DGV.RowPrePaint
        Call FORMATO_FILAS_DATAGRIDVIEW()
    End Sub
    Private Sub FORMATO_FILAS_DATAGRIDVIEW()
        Dim style As New DataGridViewCellStyle
        style.Font = New Font(DGV.Font, FontStyle.Bold)
        For Each Row As DataGridViewRow In DGV.Rows
            If Row.Cells(31).Value = False And Row.Cells(32).Value = False Then
                Row.DefaultCellStyle.ForeColor = Color.Blue
            End If
        Next
    End Sub
    Private Sub DGV_Click(sender As Object, e As EventArgs) Handles DGV.Click
        If Me.DGV.RowCount > 0 Then
            For I = 0 To Me.DGV.RowCount - 1
                If DGV.Rows(I).Selected = True Then
                    Dim FILA = DGV.CurrentRow.Index
                    If Not IsDBNull(Me.DGV.Item(0, FILA).Value) Then : hmID_HIST_M = Me.DGV.Item(0, FILA).Value : Else : hmID_HIST_M = 0 : End If
                    If Not IsDBNull(Me.DGV.Item(1, I).Value) Then : hmFEC_ING_PAME = Me.DGV.Item(1, I).Value : Else : hmFEC_ING_PAME = "" : End If
                    If Not IsDBNull(Me.DGV.Item(2, I).Value) Then : hmID_M_P = Me.DGV.Item(2, I).Value : Else : hmID_M_P = 0 : End If
                    If Not IsDBNull(Me.DGV.Item(3, I).Value) Then : hmCODIGO = Me.DGV.Item(3, I).Value : Else : hmCODIGO = "" : End If
                    If Not IsDBNull(Me.DGV.Item(4, I).Value) Then : hmAN = Me.DGV.Item(4, I).Value : Else : hmAN = "" : End If
                    If Not IsDBNull(Me.DGV.Item(5, I).Value) Then : hmFEC_MOV = Me.DGV.Item(5, I).Value : Else : hmFEC_MOV = "" : End If
                    If Not IsDBNull(Me.DGV.Item(6, I).Value) Then : hmFEC_APLICA = Me.DGV.Item(6, I).Value : Else : hmFEC_APLICA = "" : End If
                    If Not IsDBNull(Me.DGV.Item(7, I).Value) Then : hmID_T_M_EXP = Me.DGV.Item(7, I).Value : Else : hmID_T_M_EXP = 0 : End If
                    If Not IsDBNull(Me.DGV.Item(8, I).Value) Then : hmD_T_M_EXP = Me.DGV.Item(8, I).Value : Else : hmD_T_M_EXP = "" : End If
                    If Not IsDBNull(Me.DGV.Item(9, I).Value) Then : hmID_ST_M_L_EXP = Me.DGV.Item(9, I).Value : Else : hmID_ST_M_L_EXP = 0 : End If
                    If Not IsDBNull(Me.DGV.Item(10, I).Value) Then : hmD_ST_M_L_EXP = Me.DGV.Item(10, I).Value : Else : hmD_ST_M_L_EXP = "" : End If
                    If Not IsDBNull(Me.DGV.Item(11, I).Value) Then : hmID_ST_M_R_EXP = Me.DGV.Item(11, I).Value : Else : hmID_ST_M_R_EXP = 0 : End If
                    If Not IsDBNull(Me.DGV.Item(12, I).Value) Then : hmD_ST_M_R_EXP = Me.DGV.Item(12, I).Value : Else : hmD_ST_M_R_EXP = "" : End If
                    If Not IsDBNull(Me.DGV.Item(13, I).Value) Then : hmD_CARGO_ES = Me.DGV.Item(13, I).Value : Else : hmD_CARGO_ES = "" : End If
                    If Not IsDBNull(Me.DGV.Item(14, I).Value) Then : hmD_GP = Me.DGV.Item(14, I).Value : Else : hmD_GP = "" : End If
                    If Not IsDBNull(Me.DGV.Item(15, I).Value) Then : hmD_GR = Me.DGV.Item(15, I).Value : Else : hmD_GR = "" : End If
                    If Not IsDBNull(Me.DGV.Item(16, I).Value) Then : hmD_N1 = Me.DGV.Item(16, I).Value : Else : hmD_N1 = "" : End If
                    If Not IsDBNull(Me.DGV.Item(17, I).Value) Then : hmD_N2 = Me.DGV.Item(17, I).Value : Else : hmD_N2 = "" : End If
                    If Not IsDBNull(Me.DGV.Item(18, I).Value) Then : hmD_N3 = Me.DGV.Item(18, I).Value : Else : hmD_N3 = "" : End If
                    If Not IsDBNull(Me.DGV.Item(19, I).Value) Then : hmD_N4 = Me.DGV.Item(19, I).Value : Else : hmD_N4 = "" : End If
                    If Not IsDBNull(Me.DGV.Item(20, I).Value) Then : hmD_N5 = Me.DGV.Item(20, I).Value : Else : hmD_N5 = "" : End If
                    If Not IsDBNull(Me.DGV.Item(21, I).Value) Then : hmD_N6 = Me.DGV.Item(21, I).Value : Else : hmD_N6 = "" : End If
                    If Not IsDBNull(Me.DGV.Item(22, I).Value) Then : hmD_N7 = Me.DGV.Item(22, I).Value : Else : hmD_N7 = "" : End If
                    If Not IsDBNull(Me.DGV.Item(23, I).Value) Then : hmD_N8 = Me.DGV.Item(23, I).Value : Else : hmD_N8 = "" : End If
                    If Not IsDBNull(Me.DGV.Item(24, I).Value) Then : hmN_ORDEN = Me.DGV.Item(24, I).Value : Else : hmN_ORDEN = "" : End If
                    If Not IsDBNull(Me.DGV.Item(25, I).Value) Then : hmD_NOMINA = Me.DGV.Item(25, I).Value : Else : hmD_NOMINA = "" : End If
                    If Not IsDBNull(Me.DGV.Item(26, I).Value) Then : hmSALARIO = Me.DGV.Item(26, I).Value : Else : hmSALARIO = "" : End If
                    If Not IsDBNull(Me.DGV.Item(27, I).Value) Then : hmDETALLE = Me.DGV.Item(27, I).Value : Else : hmDETALLE = "" : End If
                    If Not IsDBNull(Me.DGV.Item(28, I).Value) Then : hmMIXTA = Me.DGV.Item(28, I).Value : Else : hmMIXTA = False : End If
                    If Not IsDBNull(Me.DGV.Item(29, I).Value) Then : hmMILITAR = Me.DGV.Item(29, I).Value : Else : hmMILITAR = False : End If
                    If Not IsDBNull(Me.DGV.Item(30, I).Value) Then : hmPAME = Me.DGV.Item(30, I).Value : Else : hmPAME = False : End If
                    If Not IsDBNull(Me.DGV.Item(31, I).Value) Then : hmAPLICAR = Me.DGV.Item(31, I).Value : Else : hmAPLICAR = False : End If
                    If Not IsDBNull(Me.DGV.Item(32, I).Value) Then : hmESTADISTICO = Me.DGV.Item(32, I).Value : Else : hmESTADISTICO = False : End If
                    If Not IsDBNull(Me.DGV.Item(33, I).Value) Then : hmUSUARIO_ACT = Me.DGV.Item(33, I).Value : Else : hmUSUARIO_ACT = "" : End If
                    If Not IsDBNull(Me.DGV.Item(34, I).Value) Then : hmFECHA_ACT = Me.DGV.Item(34, I).Value : Else : hmFECHA_ACT = "" : End If
                    If Not IsDBNull(Me.DGV.Item(35, I).Value) Then : hmD_ADMIN = Me.DGV.Item(35, I).Value : Else : hmD_ADMIN = "" : End If
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub Frm_Movimientos_Exp_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        MOVIMIENTOS = False
    End Sub
    Private Sub Frm_Movimientos_Exp_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        MOVIMIENTOS = False
    End Sub
    Private Sub Frm_Movimientos_Exp_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        MOVIMIENTOS = False
    End Sub
    Private Sub Frm_Movimientos_Exp_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        MOVIMIENTOS = False
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call VERIFICAR_ACCESOS("060") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Movimientos_Exp_Buscar.ShowDialog()
        If CERRAR = False Then
            Me.Cursor = Cursors.WaitCursor
            RemoveHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
            RemoveHandler Me.DateTimePicker1.ValueChanged, AddressOf DateTimePicker1_ValueChanged
            RemoveHandler Me.DateTimePicker2.ValueChanged, AddressOf DateTimePicker2_ValueChanged
            Me.DateTimePicker1.Value = FECHA01
            FECHA01 = Me.DateTimePicker1.Text
            FECHA1 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA01 + "', 105), 23)"
            Me.DateTimePicker2.Value = FECHA02
            FECHA02 = Me.DateTimePicker2.Text
            FECHA2 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA02 + "', 105), 23)"
            Me.ComboBox1.Text = cD_T_M_EXP
            If isuID_TIPO_USUARIO = 5 Then
                CADENAsql = "SELECT * FROM [dbo].[VISTA HISTORICO DE MOVIMIENTO] WHERE FEC_MOV BETWEEN " & FECHA1 & " AND " & FECHA2 & " AND ID_T_M_EXP = " & Me.ComboBox1.SelectedValue & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY FEC_MOV, ID_M_P"
            Else
                CADENAsql = "SELECT * FROM [dbo].[VISTA HISTORICO DE MOVIMIENTO] WHERE FEC_MOV BETWEEN " & FECHA1 & " AND " & FECHA2 & " AND ID_T_M_EXP = " & Me.ComboBox1.SelectedValue & " ORDER BY FEC_MOV, ID_M_P"
            End If
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
            AddHandler Me.DateTimePicker1.ValueChanged, AddressOf DateTimePicker1_ValueChanged
            AddHandler Me.DateTimePicker2.ValueChanged, AddressOf DateTimePicker2_ValueChanged
            AddHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
            Me.Cursor = Cursors.Default
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call VERIFICAR_ACCESOS("061") : If HAY_ACCESO = False Then : Exit Sub : End If
        If Me.ComboBox1.Text = "CAMBIO DE DATOS" Or Me.ComboBox1.Text = "CAMBIO SALARIAL" Then
            MsgBox("No se pueden generar movimientos de Tipo <CAMBIO DE DATOS> o <CAMBIO SALARIAL>", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        Frm_Movimientos_Exp_Add.ShowDialog()
        If CERRAR = False Then
            FECHA01 = Me.DateTimePicker1.Text
            FECHA1 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA01 + "', 105), 23)"
            FECHA02 = Me.DateTimePicker2.Text
            FECHA2 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA02 + "', 105), 23)"
            If isuID_TIPO_USUARIO = 5 Then
                CADENAsql = "SELECT * FROM [dbo].[VISTA HISTORICO DE MOVIMIENTO] WHERE FEC_MOV BETWEEN " & FECHA1 & " AND " & FECHA2 & " AND ID_T_M_EXP = " & Me.ComboBox1.SelectedValue & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY FEC_MOV, ID_M_P"
            Else
                CADENAsql = "SELECT * FROM [dbo].[VISTA HISTORICO DE MOVIMIENTO] WHERE FEC_MOV BETWEEN " & FECHA1 & " AND " & FECHA2 & " AND ID_T_M_EXP = " & Me.ComboBox1.SelectedValue & " ORDER BY FEC_MOV, ID_M_P"
            End If
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call VERIFICAR_ACCESOS("063") : If HAY_ACCESO = False Then : Exit Sub : End If
        If hmID_HIST_M = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If Me.ComboBox1.Text = "CAMBIO DE DATOS" Or Me.ComboBox1.Text = "CAMBIO SALARIAL" Then
            MsgBox("No se pueden generar movimientos de Tipo <CAMBIO DE DATOS> o <CAMBIO SALARIAL>", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        Frm_Movimientos_Exp_Upd.ShowDialog()
        If CERRAR = False Then
            FECHA01 = Me.DateTimePicker1.Text
            FECHA1 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA01 + "', 105), 23)"
            FECHA02 = Me.DateTimePicker2.Text
            FECHA2 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA02 + "', 105), 23)"
            If isuID_TIPO_USUARIO = 5 Then
                CADENAsql = "SELECT * FROM [dbo].[VISTA HISTORICO DE MOVIMIENTO] WHERE FEC_MOV BETWEEN " & FECHA1 & " AND " & FECHA2 & " AND ID_T_M_EXP = " & Me.ComboBox1.SelectedValue & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY FEC_MOV, ID_M_P"
            Else
                CADENAsql = "SELECT * FROM [dbo].[VISTA HISTORICO DE MOVIMIENTO] WHERE FEC_MOV BETWEEN " & FECHA1 & " AND " & FECHA2 & " AND ID_T_M_EXP = " & Me.ComboBox1.SelectedValue & " ORDER BY FEC_MOV, ID_M_P"
            End If
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        End If
    End Sub
    Private Sub DGV_DoubleClick(sender As Object, e As EventArgs) Handles DGV.DoubleClick
        Call VERIFICAR_ACCESOS("063") : If HAY_ACCESO = False Then : Exit Sub : End If
        If hmID_HIST_M = 0 Then
            MsgBox("Debe seleccionar el registro que desea Editar", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If Me.ComboBox1.Text = "CAMBIO DE DATOS" Or Me.ComboBox1.Text = "CAMBIO SALARIAL" Then
            MsgBox("No se pueden generar movimientos de Tipo <CAMBIO DE DATOS> o <CAMBIO SALARIAL>", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        Frm_Movimientos_Exp_Upd.ShowDialog()
        If CERRAR = False Then
            FECHA01 = Me.DateTimePicker1.Text
            FECHA1 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA01 + "', 105), 23)"
            FECHA02 = Me.DateTimePicker2.Text
            FECHA2 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA02 + "', 105), 23)"
            If isuID_TIPO_USUARIO = 5 Then
                CADENAsql = "SELECT * FROM [dbo].[VISTA HISTORICO DE MOVIMIENTO] WHERE FEC_MOV BETWEEN " & FECHA1 & " AND " & FECHA2 & " AND ID_T_M_EXP = " & Me.ComboBox1.SelectedValue & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY FEC_MOV, ID_M_P"
            Else
                CADENAsql = "SELECT * FROM [dbo].[VISTA HISTORICO DE MOVIMIENTO] WHERE FEC_MOV BETWEEN " & FECHA1 & " AND " & FECHA2 & " AND ID_T_M_EXP = " & Me.ComboBox1.SelectedValue & " ORDER BY FEC_MOV, ID_M_P"
            End If
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        End If
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Call VERIFICAR_ACCESOS("064") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Movimientos_Exp_Imprimir.ShowDialog()
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Call VERIFICAR_ACCESOS("062") : If HAY_ACCESO = False Then : Exit Sub : End If
        If hmID_HIST_M = 0 Then
            MsgBox("Debe seleccionar el registro que desea Eliminar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        If hmAPLICAR = True Then
            MsgBox("El Registro seleccionado ya se encuentra Aplicado y no puede ser Eliminado", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Esta seguro de Eliminar el registro seleccionado?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            Call ELIMINAR_REGISTRO_DE_MOVIMIENTO()
            FECHA01 = Me.DateTimePicker1.Text
            FECHA1 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA01 + "', 105), 23)"
            FECHA02 = Me.DateTimePicker2.Text
            FECHA2 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA02 + "', 105), 23)"
            If isuID_TIPO_USUARIO = 5 Then
                CADENAsql = "SELECT * FROM [dbo].[VISTA HISTORICO DE MOVIMIENTO] WHERE FEC_MOV BETWEEN " & FECHA1 & " AND " & FECHA2 & " AND ID_T_M_EXP = " & Me.ComboBox1.SelectedValue & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY FEC_MOV, ID_M_P"
            Else
                CADENAsql = "SELECT * FROM [dbo].[VISTA HISTORICO DE MOVIMIENTO] WHERE FEC_MOV BETWEEN " & FECHA1 & " AND " & FECHA2 & " AND ID_T_M_EXP = " & Me.ComboBox1.SelectedValue & " ORDER BY FEC_MOV, ID_M_P"
            End If
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        End If
    End Sub
    Private Sub ELIMINAR_REGISTRO_DE_MOVIMIENTO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "DELETE FROM [HISTORICO DE MOVIMIENTO] WHERE ID_HIST_M = " & CInt(hmID_HIST_M) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub

End Class