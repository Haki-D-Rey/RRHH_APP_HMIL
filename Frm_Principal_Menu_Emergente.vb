Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Principal_Menu_Emergente
    Private Sub Frm_Principal_Menu_Emergente_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Frm_Principal.Timer1.Start()
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Principal_Menu_Emergente_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim r As Rectangle = My.Computer.Screen.WorkingArea
        Location = New Point(r.Width - Width, r.Height - Height)
        Call CARGAR_MOVIMIENTOS_A_APLICAR()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Frm_Principal.Timer1.Start()
        Me.Close()
    End Sub
    Private Sub CARGAR_MOVIMIENTOS_A_APLICAR()
        Call Clases.BUSCAR_FECHA_Y_HORA()
        FECHA01 = Mid(FECHA_SERVIDOR, 1, 10)
        Dim FECHA1 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA01 + "', 105), 23)"
        CADENAsql = "SELECT * FROM [dbo].[VISTA HISTORICO DE MOVIMIENTO] WHERE (FEC_APLICA > = " & FECHA1 & ") AND (ID_T_M_EXP = 1) AND (APLICAR = 'FALSE' AND ESTADISTICO = 'FALSE') ORDER BY FEC_APLICA"
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
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
        Me.Label1.Text = DGV.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW()
        Me.DGV.Columns(0).HeaderText = "Id"
        Me.DGV.Columns(0).Width = 0
        Me.DGV.Columns(0).Visible = False
        Me.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(1).HeaderText = "Fecha Ing. PAME"
        Me.DGV.Columns(1).Width = 0
        Me.DGV.Columns(1).Visible = False
        Me.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(2).HeaderText = "ID_M_P"
        Me.DGV.Columns(2).Width = 0
        Me.DGV.Columns(2).Visible = False

        Me.DGV.Columns(3).HeaderText = "Codigo"
        Me.DGV.Columns(3).Width = 70
        Me.DGV.Columns(3).Visible = True
        Me.DGV.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(4).HeaderText = "Apellidos y Nombres"
        Me.DGV.Columns(4).Width = 240
        Me.DGV.Columns(4).Visible = True
        Me.DGV.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(5).HeaderText = "Fecha Grabacion"
        Me.DGV.Columns(5).Width = 0
        Me.DGV.Columns(5).Visible = False
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
        Me.DGV.Columns(10).Width = 0
        Me.DGV.Columns(10).Visible = False
        Me.DGV.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(11).HeaderText = "ID_ST_M_R_EXP"
        Me.DGV.Columns(11).Width = 0
        Me.DGV.Columns(11).Visible = False

        Me.DGV.Columns(12).HeaderText = "Causa Real"
        Me.DGV.Columns(12).Width = 0
        Me.DGV.Columns(12).Visible = False
        Me.DGV.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(13).HeaderText = "Cargo"
        Me.DGV.Columns(13).Width = 0
        Me.DGV.Columns(13).Visible = False
        Me.DGV.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(14).HeaderText = "Gp"
        Me.DGV.Columns(14).Width = 0
        Me.DGV.Columns(14).Visible = False
        Me.DGV.Columns(14).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(15).HeaderText = "Gr"
        Me.DGV.Columns(15).Width = 0
        Me.DGV.Columns(15).Visible = False
        Me.DGV.Columns(15).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(16).HeaderText = "Nivel 1"
        Me.DGV.Columns(16).Width = 0
        Me.DGV.Columns(16).Visible = False
        Me.DGV.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(17).HeaderText = "Nivel 2"
        Me.DGV.Columns(17).Width = 0
        Me.DGV.Columns(17).Visible = False
        Me.DGV.Columns(17).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(18).HeaderText = "Nivel 3"
        Me.DGV.Columns(18).Width = 0
        Me.DGV.Columns(18).Visible = False
        Me.DGV.Columns(18).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(19).HeaderText = "Nivel 4"
        Me.DGV.Columns(19).Width = 0
        Me.DGV.Columns(19).Visible = False
        Me.DGV.Columns(19).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(20).HeaderText = "Nivel 5"
        Me.DGV.Columns(20).Width = 0
        Me.DGV.Columns(20).Visible = False
        Me.DGV.Columns(20).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(21).HeaderText = "Nivel 6"
        Me.DGV.Columns(21).Width = 0
        Me.DGV.Columns(21).Visible = False
        Me.DGV.Columns(21).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(22).HeaderText = "Nivel 7"
        Me.DGV.Columns(22).Width = 0
        Me.DGV.Columns(22).Visible = False
        Me.DGV.Columns(22).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(23).HeaderText = "Nivel 8"
        Me.DGV.Columns(23).Width = 0
        Me.DGV.Columns(23).Visible = False
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            TITULO_EXCEL = "PROYECCION DE MOVIMIENTOS"
            EXPORTAR_DATOS_A_EXCEL(DGV, "Movimientos Proyectados a Aplicar (BAJAS)")
        Catch ex As Exception
            MessageBox.Show("No se puede generar el documento Excel.", "Mensaje del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class