Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Expedientes_Activos_Jefe_Inmediato
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.TextBox2.SelectAll()
        Me.TextBox2.Focus()
        CERRAR = True
        Me.Close()
    End Sub

    Private Sub Frm_Expedientes_Activos_Jefe_Inmediato_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox2.SelectAll()
            Me.TextBox2.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Expedientes_Activos_Jefe_Inmediato_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = "0"
        Me.TextBox2.Text = Frm_Expedientes_Activos.TextBox44.Text
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
        Me.TextBox2.SelectAll()
        Me.TextBox2.Focus()
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If Len(Me.TextBox2.Text) = 0 Then
            MsgBox("Debe digitar correctamente el Texto que desea Buscar", vbInformation, "Mensaje del sistema")
            Me.TextBox2.Focus()
        Else
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Me.TextBox2.SelectAll()
            Me.TextBox2.Focus()
        End If
    End Sub
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If (e.KeyCode = Keys.Enter) Then
            If Len(Me.TextBox2.Text) = 0 Then
                MsgBox("Debe digitar correctamente el Texto que desea Buscar", vbInformation, "Mensaje del sistema")
                Me.TextBox2.Focus()
            Else
                Call CARGAR_DATAGRIDVIEW()
                Call ARMAR_DATAGRIDVIEW()
                Me.TextBox2.SelectAll()
                Me.TextBox2.Focus()
            End If
        End If
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW()
        CADENAsql = "Select * from [VISTA MAESTRO DE CARGOS] WHERE ID_SITUACION = '1' AND (AN LIKE '%" & Me.TextBox2.Text & "%') ORDER BY AN"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE CARGOS]")
        DGV.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE CARGOS]")
        Me.Label2.Text = "Total de Registros: " & DGV.RowCount
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW()
        Me.DGV.Columns(0).HeaderText = "ID_M_C"
        Me.DGV.Columns(0).Width = 0
        Me.DGV.Columns(0).Visible = False
        Me.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(1).HeaderText = "ORDEN"
        Me.DGV.Columns(1).Width = 0
        Me.DGV.Columns(1).Visible = False

        Me.DGV.Columns(2).HeaderText = "ID_N1"
        Me.DGV.Columns(2).Width = 0
        Me.DGV.Columns(2).Visible = False

        Me.DGV.Columns(3).HeaderText = "ORDEN_N1"
        Me.DGV.Columns(3).Width = 0
        Me.DGV.Columns(3).Visible = False

        Me.DGV.Columns(4).HeaderText = "D_N1"
        Me.DGV.Columns(4).Width = 0
        Me.DGV.Columns(4).Visible = False

        Me.DGV.Columns(5).HeaderText = "ID_N2"
        Me.DGV.Columns(5).Width = 0
        Me.DGV.Columns(5).Visible = False

        Me.DGV.Columns(6).HeaderText = "ORDEN_N2"
        Me.DGV.Columns(6).Width = 0
        Me.DGV.Columns(6).Visible = False

        Me.DGV.Columns(7).HeaderText = "D_N2"
        Me.DGV.Columns(7).Width = 0
        Me.DGV.Columns(7).Visible = False

        Me.DGV.Columns(8).HeaderText = "ID_N3"
        Me.DGV.Columns(8).Width = 0
        Me.DGV.Columns(8).Visible = False

        Me.DGV.Columns(9).HeaderText = "ORDEN_N3"
        Me.DGV.Columns(9).Width = 0
        Me.DGV.Columns(9).Visible = False

        Me.DGV.Columns(10).HeaderText = "D_N3"
        Me.DGV.Columns(10).Width = 0
        Me.DGV.Columns(10).Visible = False

        Me.DGV.Columns(11).HeaderText = "ID_N4"
        Me.DGV.Columns(11).Width = 0
        Me.DGV.Columns(11).Visible = False

        Me.DGV.Columns(12).HeaderText = "ORDEN_N4"
        Me.DGV.Columns(12).Width = 0
        Me.DGV.Columns(12).Visible = False

        Me.DGV.Columns(13).HeaderText = "D_N4"
        Me.DGV.Columns(13).Width = 0
        Me.DGV.Columns(13).Visible = False

        Me.DGV.Columns(14).HeaderText = "ID_N5"
        Me.DGV.Columns(14).Width = 0
        Me.DGV.Columns(14).Visible = False

        Me.DGV.Columns(15).HeaderText = "ORDEN_N5"
        Me.DGV.Columns(15).Width = 0
        Me.DGV.Columns(15).Visible = False

        Me.DGV.Columns(16).HeaderText = "D_N5"
        Me.DGV.Columns(16).Width = 0
        Me.DGV.Columns(16).Visible = False

        Me.DGV.Columns(17).HeaderText = "ID_N6"
        Me.DGV.Columns(17).Width = 0
        Me.DGV.Columns(17).Visible = False

        Me.DGV.Columns(18).HeaderText = "ORDEN_N6"
        Me.DGV.Columns(18).Width = 0
        Me.DGV.Columns(18).Visible = False

        Me.DGV.Columns(19).HeaderText = "D_N6"
        Me.DGV.Columns(19).Width = 0
        Me.DGV.Columns(19).Visible = False

        Me.DGV.Columns(20).HeaderText = "ID_N7"
        Me.DGV.Columns(20).Width = 0
        Me.DGV.Columns(20).Visible = False

        Me.DGV.Columns(21).HeaderText = "ORDEN_N7"
        Me.DGV.Columns(21).Width = 0
        Me.DGV.Columns(21).Visible = False

        Me.DGV.Columns(22).HeaderText = "D_N7"
        Me.DGV.Columns(22).Width = 0
        Me.DGV.Columns(22).Visible = False

        Me.DGV.Columns(23).HeaderText = "ID_N8"
        Me.DGV.Columns(23).Width = 0
        Me.DGV.Columns(23).Visible = False

        Me.DGV.Columns(24).HeaderText = "ORDEN_N8"
        Me.DGV.Columns(24).Width = 0
        Me.DGV.Columns(24).Visible = False

        Me.DGV.Columns(25).HeaderText = "D_N8"
        Me.DGV.Columns(25).Width = 0
        Me.DGV.Columns(25).Visible = False

        Me.DGV.Columns(26).HeaderText = "ID_T_ES"
        Me.DGV.Columns(26).Width = 0
        Me.DGV.Columns(26).Visible = False

        Me.DGV.Columns(27).HeaderText = "D_T_ES"
        Me.DGV.Columns(27).Width = 0
        Me.DGV.Columns(27).Visible = False

        Me.DGV.Columns(28).HeaderText = "Orden [Mixto]"
        Me.DGV.Columns(28).Width = 50
        Me.DGV.Columns(28).Visible = True
        Me.DGV.Columns(28).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(28).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(29).HeaderText = "N_ORDEN_MILITAR"
        Me.DGV.Columns(29).Width = 0
        Me.DGV.Columns(29).Visible = False

        Me.DGV.Columns(30).HeaderText = "N_ORDEN_PAME"
        Me.DGV.Columns(30).Width = 0
        Me.DGV.Columns(30).Visible = False

        Me.DGV.Columns(31).HeaderText = "ID_CARGO_ES"
        Me.DGV.Columns(31).Width = 0
        Me.DGV.Columns(31).Visible = False

        Me.DGV.Columns(32).HeaderText = "Cargo"
        Me.DGV.Columns(32).Width = 150
        Me.DGV.Columns(32).Visible = True
        Me.DGV.Columns(32).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        'Me.DGV.Columns(32).DefaultCellStyle.BackColor = Color.LightGreen

        Me.DGV.Columns(33).HeaderText = "ID_GP"
        Me.DGV.Columns(33).Width = 0
        Me.DGV.Columns(33).Visible = False

        Me.DGV.Columns(34).HeaderText = "D_GP"
        Me.DGV.Columns(34).Width = 0
        Me.DGV.Columns(34).Visible = False

        Me.DGV.Columns(35).HeaderText = "ID_GR"
        Me.DGV.Columns(35).Width = 0
        Me.DGV.Columns(35).Visible = False

        Me.DGV.Columns(36).HeaderText = "Grado Real"
        Me.DGV.Columns(36).Width = 80
        Me.DGV.Columns(36).Visible = True
        Me.DGV.Columns(36).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(37).HeaderText = "ID_CAT_SALARIAL"
        Me.DGV.Columns(37).Width = 0
        Me.DGV.Columns(37).Visible = False

        Me.DGV.Columns(38).HeaderText = "D_CAT_SALARIAL"
        Me.DGV.Columns(38).Width = 0
        Me.DGV.Columns(38).Visible = False

        Me.DGV.Columns(39).HeaderText = "ID_M_P"
        Me.DGV.Columns(39).Width = 0
        Me.DGV.Columns(39).Visible = False

        Me.DGV.Columns(40).HeaderText = "Codigo"
        Me.DGV.Columns(40).Width = 80
        Me.DGV.Columns(40).Visible = True
        Me.DGV.Columns(40).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(41).HeaderText = "Apellidos y Empleados"
        Me.DGV.Columns(41).Width = 190
        Me.DGV.Columns(41).Visible = True
        Me.DGV.Columns(41).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(42).HeaderText = "ID_SITUACION"
        Me.DGV.Columns(42).Width = 0
        Me.DGV.Columns(42).Visible = False

        Me.DGV.Columns(43).HeaderText = "Jefe"
        Me.DGV.Columns(43).Width = 0
        Me.DGV.Columns(43).Visible = False
        Me.DGV.Columns(43).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(44).HeaderText = "Mixta"
        Me.DGV.Columns(44).Width = 0
        Me.DGV.Columns(44).Visible = False
        Me.DGV.Columns(44).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(45).HeaderText = "Militar"
        Me.DGV.Columns(45).Width = 0
        Me.DGV.Columns(45).Visible = False
        Me.DGV.Columns(45).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(46).HeaderText = "Pame"
        Me.DGV.Columns(46).Width = 0
        Me.DGV.Columns(46).Visible = False
        Me.DGV.Columns(46).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(47).HeaderText = "ID_ESTABLECIMIENTO"
        Me.DGV.Columns(47).Width = 0
        Me.DGV.Columns(47).Visible = False

        Me.DGV.Columns(48).HeaderText = "D_ESTABLECIMIENTO"
        Me.DGV.Columns(48).Width = 0
        Me.DGV.Columns(48).Visible = False

        Me.DGV.Columns(49).HeaderText = "D_ADMIN"
        Me.DGV.Columns(49).Width = 0
        Me.DGV.Columns(49).Visible = False
    End Sub
    Private Sub DGV_Click(sender As Object, e As EventArgs) Handles DGV.Click
        If Me.DGV.RowCount > 0 Then
            For I = 0 To Me.DGV.RowCount - 1
                If DGV.Rows(I).Selected = True Then
                    'vID_M_P = Me.DGV.Item(39, I).Value
                    'vID_M_C = Me.DGV.Item(0, I).Value
                    bEMPLEADO = Me.DGV.Item(41, I).Value
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub DGV_DoubleClick(sender As Object, e As EventArgs) Handles DGV.DoubleClick
        If bEMPLEADO <> "" Then
            Me.TextBox1.Text = "0"
            Me.TextBox2.Text = "Buscar Empleado..."
            Me.TextBox2.SelectAll()
            Me.TextBox2.Focus()
            CERRAR = False
            Me.Close()
        Else
            MsgBox("Debe seleccionar el registro que desea Visualizar", vbInformation, "Mensaje del Sistema")
        End If
    End Sub
End Class