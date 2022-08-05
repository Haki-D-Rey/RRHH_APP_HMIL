Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Personal_Tercerizado_Buscar
    Private Sub Frm_Personal_Tercerizado_Buscar_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = "0"
            Me.TextBox2.Text = "Buscar ..."
            Me.TextBox2.SelectAll()
            Me.TextBox2.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Personal_Tercerizado_Buscar_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = "0"
        Me.TextBox2.Text = "Buscar ..."
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
        Me.TextBox2.SelectAll()
        Me.TextBox2.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox1.Text = "0"
        Me.TextBox2.Text = "Buscar ..."
        Me.TextBox2.SelectAll()
        Me.TextBox2.Focus()
        CERRAR = True
        Me.Close()
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
        CADENAsql = "Select ID_P_T, ID_E_T, D_E_T, CODIGO, APELLIDOS, NOMBRES, AN, N_CEDULA, ID_T_SEXO, D_T_SEXO, TELEFONO_1, TELEFONO_2, D_CARGO_ES, ID_ESTADO_P, D_ESTADO_P, DIRECCION, NSS, OBSERVACIONES from [VISTA MAESTRO DE PERSONAL TERCERIZADO] WHERE AN LIKE '%" & Trim(Me.TextBox2.Text) & "%' ORDER BY D_E_T, AN"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE PERSONAL TERCERIZADO]")
        DGV.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE PERSONAL TERCERIZADO]")
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

        Me.DGV.Columns(1).HeaderText = "ID_E_T"
        Me.DGV.Columns(1).Width = 0
        Me.DGV.Columns(1).Visible = False

        Me.DGV.Columns(2).HeaderText = "Empresa"
        Me.DGV.Columns(2).Width = 130
        Me.DGV.Columns(2).Visible = True
        Me.DGV.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(3).HeaderText = "Codigo"
        Me.DGV.Columns(3).Width = 100
        Me.DGV.Columns(3).Visible = True
        Me.DGV.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(4).HeaderText = "APELLIDOS"
        Me.DGV.Columns(4).Width = 0
        Me.DGV.Columns(4).Visible = False

        Me.DGV.Columns(5).HeaderText = "NOMBRES"
        Me.DGV.Columns(5).Width = 0
        Me.DGV.Columns(5).Visible = False

        Me.DGV.Columns(6).HeaderText = "Apellidos y Nombres"
        Me.DGV.Columns(6).Width = 300
        Me.DGV.Columns(6).Visible = True
        Me.DGV.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(7).HeaderText = "Cedula"
        Me.DGV.Columns(7).Width = 110
        Me.DGV.Columns(7).Visible = True
        Me.DGV.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(8).HeaderText = "ID_T_SEXO"
        Me.DGV.Columns(8).Width = 0
        Me.DGV.Columns(8).Visible = False

        Me.DGV.Columns(9).HeaderText = "Sexo"
        Me.DGV.Columns(9).Width = 0
        Me.DGV.Columns(9).Visible = False
        Me.DGV.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(10).HeaderText = "Telef. (C)"
        Me.DGV.Columns(10).Width = 0
        Me.DGV.Columns(10).Visible = False
        Me.DGV.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(11).HeaderText = "Telef. (M)"
        Me.DGV.Columns(11).Width = 0
        Me.DGV.Columns(11).Visible = False
        Me.DGV.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(12).HeaderText = "Cargo"
        Me.DGV.Columns(12).Width = 0
        Me.DGV.Columns(12).Visible = False
        Me.DGV.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(13).HeaderText = "ID_ESTADO_P"
        Me.DGV.Columns(13).Width = 0
        Me.DGV.Columns(13).Visible = False

        Me.DGV.Columns(14).HeaderText = "Estado"
        Me.DGV.Columns(14).Width = 0
        Me.DGV.Columns(14).Visible = False
        Me.DGV.Columns(14).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(15).HeaderText = "DIRECCION"
        Me.DGV.Columns(15).Width = 0
        Me.DGV.Columns(15).Visible = False

        Me.DGV.Columns(16).HeaderText = "NSS"
        Me.DGV.Columns(16).Width = 0
        Me.DGV.Columns(16).Visible = False

        Me.DGV.Columns(17).HeaderText = "OBSERVACIONES"
        Me.DGV.Columns(17).Width = 0
        Me.DGV.Columns(17).Visible = False

    End Sub
    Private Sub DGV_Click(sender As Object, e As EventArgs) Handles DGV.Click
        If Me.DGV.RowCount > 0 Then
            For I = 0 To Me.DGV.RowCount - 1
                If DGV.Rows(I).Selected = True Then
                    ptID_P_T = Me.DGV.Item(0, I).Value
                    ptD_E_T = Me.DGV.Item(2, I).Value
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub DGV_DoubleClick(sender As Object, e As EventArgs) Handles DGV.DoubleClick
        If ptID_P_T <> 0 Then
            Me.TextBox1.Text = "0"
            Me.TextBox2.Text = "Buscar ..."
            Me.TextBox2.SelectAll()
            Me.TextBox2.Focus()
            CERRAR = False
            Me.Close()
        Else
            MsgBox("Debe seleccionar el registro que desea Visualizar", vbInformation, "Mensaje del Sistema")
        End If
    End Sub
    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs)
        Me.TextBox2.SelectAll()
        Me.TextBox2.Focus()
    End Sub
End Class