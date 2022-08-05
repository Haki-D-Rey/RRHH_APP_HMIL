Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Movimientos_Exp_Buscar
    Private Sub Frm_Movimientos_Exp_Buscar_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox2.SelectAll()
            Me.TextBox2.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Movimientos_Exp_Buscar_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = "0"
        Me.TextBox2.Text = "Buscar..."
        Me.RadioButton2.Checked = True
        Me.CheckBox1.Checked = True
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
        Me.TextBox2.SelectAll()
        Me.TextBox2.Focus()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
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
        If isuID_TIPO_USUARIO = 5 Then
            If Me.RadioButton1.Checked = True Then
                CADENAsql = "Select * from [VISTA HISTORICO DE MOVIMIENTO] WHERE (CODIGO LIKE '%" & Me.TextBox2.Text & "%') AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY CODIGO, FEC_MOV"
                GoTo SALTO
            End If
            If Me.RadioButton2.Checked = True Then
                If Me.CheckBox1.Checked = False Then
                    Dim buscar As String = ""
                    Dim tamañoArreglo As Integer = 0
                    buscar = Me.TextBox2.Text.Trim
                    Dim ArrCadena As String() = buscar.Split(" ")
                    tamañoArreglo = ArrCadena.GetUpperBound(0) + 1
                    If tamañoArreglo = 0 Then
                        CADENAsql = "SELECT * FROM [dbo].[VISTA HISTORICO DE MOVIMIENTO] WHERE (AN LIKE '%" & Me.TextBox2.Text & "%') AND D_ADMIN = '" & isuUSUARIO & "'"
                    ElseIf tamañoArreglo = 1 Then
                        CADENAsql = "SELECT * FROM [dbo].[VISTA HISTORICO DE MOVIMIENTO WHERE AN LIKE '%" & buscar & "%' AND D_ADMIN = '" & isuUSUARIO & "'"
                    ElseIf tamañoArreglo > 1 Then
                        CADENAsql = "SELECT * FROM [dbo].[VISTA HISTORICO DE MOVIMIENTO] WHERE AN LIKE '%" & ArrCadena(0) & "%' AND D_ADMIN = '" & isuUSUARIO & "'"
                        For index = 0 To ArrCadena.GetUpperBound(0)
                            CADENAsql = CADENAsql & " UNION SELECT * FROM [dbo].[VISTA HISTORICO DE MOVIMIENTO] WHERE AN LIKE '%" & ArrCadena(index) & "%' AND D_ADMIN = '" & isuUSUARIO & "'"
                        Next
                    End If
                    CADENAsql = CADENAsql & "  ORDER BY CODIGO, FEC_MOV"
                    GoTo SALTO
                End If
                If Me.CheckBox1.Checked = True Then
                    CADENAsql = "Select * from [VISTA HISTORICO DE MOVIMIENTO] WHERE (AN LIKE '%" & Me.TextBox2.Text & "%') AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY CODIGO, FEC_MOV"
                End If
            End If
        Else
            If Me.RadioButton1.Checked = True Then
                CADENAsql = "Select * from [VISTA HISTORICO DE MOVIMIENTO] WHERE (CODIGO LIKE '%" & Me.TextBox2.Text & "%') ORDER BY CODIGO, FEC_MOV"
                GoTo SALTO
            End If
            If Me.RadioButton2.Checked = True Then
                If Me.CheckBox1.Checked = False Then
                    Dim buscar As String = ""
                    Dim tamañoArreglo As Integer = 0
                    buscar = Me.TextBox2.Text.Trim
                    Dim ArrCadena As String() = buscar.Split(" ")
                    tamañoArreglo = ArrCadena.GetUpperBound(0) + 1
                    If tamañoArreglo = 0 Then
                        CADENAsql = "SELECT * FROM [dbo].[VISTA HISTORICO DE MOVIMIENTO] WHERE (AN LIKE '%" & Me.TextBox2.Text & "%')"
                    ElseIf tamañoArreglo = 1 Then
                        CADENAsql = "SELECT * FROM [dbo].[VISTA HISTORICO DE MOVIMIENTO WHERE AN LIKE '%" & buscar & "%'"
                    ElseIf tamañoArreglo > 1 Then
                        CADENAsql = "SELECT * FROM [dbo].[VISTA HISTORICO DE MOVIMIENTO] WHERE AN LIKE '%" & ArrCadena(0) & "%'"
                        For index = 0 To ArrCadena.GetUpperBound(0)
                            CADENAsql = CADENAsql & " UNION SELECT * FROM [dbo].[VISTA HISTORICO DE MOVIMIENTO] WHERE AN LIKE '%" & ArrCadena(index) & "%'"
                        Next
                    End If
                    CADENAsql = CADENAsql & "  ORDER BY CODIGO, FEC_MOV"
                    GoTo SALTO
                End If
                If Me.CheckBox1.Checked = True Then
                    CADENAsql = "Select * from [VISTA HISTORICO DE MOVIMIENTO] WHERE (AN LIKE '%" & Me.TextBox2.Text & "%') ORDER BY CODIGO, FEC_MOV"
                End If
            End If
        End If
SALTO:
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA HISTORICO DE MOVIMIENTO]")
        DGV.DataSource = MiDataSet.Tables("[VISTA HISTORICO DE MOVIMIENTO]")
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
        Me.DGV.Columns(4).Width = 180
        Me.DGV.Columns(4).Visible = True
        Me.DGV.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(5).HeaderText = "Fecha Grabacion"
        Me.DGV.Columns(5).Width = 70
        Me.DGV.Columns(5).Visible = True
        Me.DGV.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(6).HeaderText = "Fecha Aplicacion"
        Me.DGV.Columns(6).Width = 0
        Me.DGV.Columns(6).Visible = False
        Me.DGV.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(7).HeaderText = "ID_T_M_EXP"
        Me.DGV.Columns(7).Width = 0
        Me.DGV.Columns(7).Visible = False

        Me.DGV.Columns(8).HeaderText = "Tipo Movimiento"
        Me.DGV.Columns(8).Width = 80
        Me.DGV.Columns(8).Visible = True
        Me.DGV.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

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
        Me.DGV.Columns(13).Width = 100
        Me.DGV.Columns(13).Visible = True
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
        Me.DGV.Columns(16).Width = 80
        Me.DGV.Columns(16).Visible = True
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
        Me.DGV.Columns(26).Width = 80
        Me.DGV.Columns(26).Visible = True
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
    Private Sub DGV_Click(sender As Object, e As EventArgs) Handles DGV.Click
        If Me.DGV.RowCount > 0 Then
            For I = 0 To Me.DGV.RowCount - 1
                If DGV.Rows(I).Selected = True Then
                    cID_HIST_M = Me.DGV.Item(0, I).Value
                    FECHA01 = "01/" & Mid(Me.DGV.Item(5, I).Value, 4, 10)
                    FECHA02 = Mid(Me.DGV.Item(5, I).Value, 1, 10)
                    cD_T_M_EXP = Me.DGV.Item(8, I).Value
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub DGV_DoubleClick(sender As Object, e As EventArgs) Handles DGV.DoubleClick
        If cID_HIST_M <> 0 Then
            Me.TextBox1.Text = "0"
            Me.TextBox2.Text = "Buscar ..."
            Me.TextBox2.SelectAll()
            Me.TextBox2.Focus()
            vID_HIST_M = cID_HIST_M
            CERRAR = False
            Me.Close()
        Else
            MsgBox("Debe seleccionar el registro que desea Visualizar", vbInformation, "Mensaje del Sistema")
        End If
    End Sub
    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If Me.RadioButton2.Checked = True Then
            Me.CheckBox1.Checked = True
            Me.TextBox2.SelectAll()
            Me.TextBox2.Focus()
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If Me.RadioButton1.Checked = True Then
            Me.CheckBox1.Checked = False
            Me.TextBox2.SelectAll()
            Me.TextBox2.Focus()
        End If
    End Sub
End Class