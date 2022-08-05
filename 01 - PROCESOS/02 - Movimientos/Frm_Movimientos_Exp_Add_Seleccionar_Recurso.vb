Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Movimientos_Exp_Add_Seleccionar_Recurso
    Private Sub Frm_Movimientos_Exp_Add_Seleccionar_Recurso_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox2.SelectAll()
            Me.TextBox2.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Movimientos_Exp_Add_Seleccionar_Recurso_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = "0"
        Me.TextBox2.Text = "Buscar..."
        Me.RadioButton1.Checked = False
        Me.RadioButton2.Checked = True
        Me.CheckBox1.Checked = True
        Me.RadioButton3.Checked = False
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
                CADENAsql = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE (CODIGO LIKE '%" & Me.TextBox2.Text & "%') AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY AN"
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
                        CADENAsql = "SELECT * FROM [dbo].[VISTA MAESTRO DE CARGOS SIT] WHERE (AN LIKE '%" & Me.TextBox2.Text & "%') AND D_ADMIN = '" & isuUSUARIO & "'"
                    ElseIf tamañoArreglo = 1 Then
                        CADENAsql = "SELECT * FROM [dbo].[VISTA MAESTRO DE CARGOS SIT] WHERE AN LIKE '%" & buscar & "%' AND D_ADMIN = '" & isuUSUARIO & "'"
                    ElseIf tamañoArreglo > 1 Then
                        CADENAsql = "SELECT * FROM [dbo].[VISTA MAESTRO DE CARGOS SIT] WHERE AN LIKE '%" & ArrCadena(0) & "%' AND D_ADMIN = '" & isuUSUARIO & "'"
                        For index = 0 To ArrCadena.GetUpperBound(0)
                            CADENAsql = CADENAsql & " UNION SELECT * FROM [dbo].[VISTA MAESTRO DE CARGOS SIT] WHERE AN LIKE '%" & ArrCadena(index) & "%' AND D_ADMIN = '" & isuUSUARIO & "'"
                        Next
                    End If
                    CADENAsql = CADENAsql & "  ORDER BY AN"
                    GoTo SALTO
                End If
                If Me.CheckBox1.Checked = True Then
                    CADENAsql = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE (AN LIKE '%" & Me.TextBox2.Text & "%') AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY AN"
                End If
            End If
            If Me.RadioButton3.Checked = True Then
                CADENAsql = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE (N_ORDEN_PAME LIKE '%" & Me.TextBox2.Text & "%') AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY AN"
                GoTo SALTO
            End If
        Else
            If Me.RadioButton1.Checked = True Then
                CADENAsql = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE (CODIGO LIKE '%" & Me.TextBox2.Text & "%') ORDER BY AN"
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
                        CADENAsql = "SELECT * FROM [dbo].[VISTA MAESTRO DE CARGOS SIT] WHERE (AN LIKE '%" & Me.TextBox2.Text & "%') "
                    ElseIf tamañoArreglo = 1 Then
                        CADENAsql = "SELECT * FROM [dbo].[VISTA MAESTRO DE CARGOS SIT] WHERE AN LIKE '%" & buscar & "%' "
                    ElseIf tamañoArreglo > 1 Then
                        CADENAsql = "SELECT * FROM [dbo].[VISTA MAESTRO DE CARGOS SIT] WHERE AN LIKE '%" & ArrCadena(0) & "%' "
                        For index = 0 To ArrCadena.GetUpperBound(0)
                            CADENAsql = CADENAsql & " UNION SELECT * FROM [dbo].[VISTA MAESTRO DE CARGOS SIT] WHERE AN LIKE '%" & ArrCadena(index) & "%' "
                        Next
                    End If
                    CADENAsql = CADENAsql & "  ORDER BY AN"
                    GoTo SALTO
                End If
                If Me.CheckBox1.Checked = True Then
                    CADENAsql = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE (AN LIKE '%" & Me.TextBox2.Text & "%') ORDER BY AN"
                End If
            End If
            If Me.RadioButton3.Checked = True Then
                CADENAsql = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE (N_ORDEN_PAME LIKE '%" & Me.TextBox2.Text & "%') ORDER BY AN"
                GoTo SALTO
            End If
        End If
SALTO:
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE CARGOS SIT]")
        DGV.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE CARGOS SIT]")
        Me.Label2.Text = DGV.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW()
        Me.DGV.Columns(0).HeaderText = "Id"
        Me.DGV.Columns(0).Width = 20
        Me.DGV.Columns(0).Visible = True
        Me.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Me.DGV.Columns(1).HeaderText = "ORDEN"
        Me.DGV.Columns(1).Width = 0
        Me.DGV.Columns(1).Visible = False

        Me.DGV.Columns(2).HeaderText = "ID_N1"
        Me.DGV.Columns(2).Width = 0
        Me.DGV.Columns(2).Visible = False

        Me.DGV.Columns(3).HeaderText = "ORDEN_N1"
        Me.DGV.Columns(3).Width = 0
        Me.DGV.Columns(3).Visible = False

        Me.DGV.Columns(4).HeaderText = "Nivel 1"
        Me.DGV.Columns(4).Width = 0
        Me.DGV.Columns(4).Visible = False
        Me.DGV.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(5).HeaderText = "ID_N2"
        Me.DGV.Columns(5).Width = 0
        Me.DGV.Columns(5).Visible = False

        Me.DGV.Columns(6).HeaderText = "ORDEN_N2"
        Me.DGV.Columns(6).Width = 0
        Me.DGV.Columns(6).Visible = False

        Me.DGV.Columns(7).HeaderText = "Nivel 2"
        Me.DGV.Columns(7).Width = 0
        Me.DGV.Columns(7).Visible = False
        Me.DGV.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(8).HeaderText = "ID_N3"
        Me.DGV.Columns(8).Width = 0
        Me.DGV.Columns(8).Visible = False

        Me.DGV.Columns(9).HeaderText = "ORDEN_N3"
        Me.DGV.Columns(9).Width = 0
        Me.DGV.Columns(9).Visible = False

        Me.DGV.Columns(10).HeaderText = "Nivel 3"
        Me.DGV.Columns(10).Width = 0
        Me.DGV.Columns(10).Visible = False
        Me.DGV.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(11).HeaderText = "ID_N4"
        Me.DGV.Columns(11).Width = 0
        Me.DGV.Columns(11).Visible = False

        Me.DGV.Columns(12).HeaderText = "ORDEN_N4"
        Me.DGV.Columns(12).Width = 0
        Me.DGV.Columns(12).Visible = False

        Me.DGV.Columns(13).HeaderText = "Nivel 4"
        Me.DGV.Columns(13).Width = 0
        Me.DGV.Columns(13).Visible = False
        Me.DGV.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(14).HeaderText = "ID_N5"
        Me.DGV.Columns(14).Width = 0
        Me.DGV.Columns(14).Visible = False

        Me.DGV.Columns(15).HeaderText = "ORDEN_N5"
        Me.DGV.Columns(15).Width = 0
        Me.DGV.Columns(15).Visible = False

        Me.DGV.Columns(16).HeaderText = "Nivel 5"
        Me.DGV.Columns(16).Width = 0
        Me.DGV.Columns(16).Visible = False
        Me.DGV.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(17).HeaderText = "ID_N6"
        Me.DGV.Columns(17).Width = 0
        Me.DGV.Columns(17).Visible = False

        Me.DGV.Columns(18).HeaderText = "ORDEN_N6"
        Me.DGV.Columns(18).Width = 0
        Me.DGV.Columns(18).Visible = False

        Me.DGV.Columns(19).HeaderText = "Nivel 6"
        Me.DGV.Columns(19).Width = 0
        Me.DGV.Columns(19).Visible = False
        Me.DGV.Columns(19).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(20).HeaderText = "ID_N7"
        Me.DGV.Columns(20).Width = 0
        Me.DGV.Columns(20).Visible = False

        Me.DGV.Columns(21).HeaderText = "ORDEN_N7"
        Me.DGV.Columns(21).Width = 0
        Me.DGV.Columns(21).Visible = False

        Me.DGV.Columns(22).HeaderText = "Nivel 7"
        Me.DGV.Columns(22).Width = 0
        Me.DGV.Columns(22).Visible = False
        Me.DGV.Columns(22).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(23).HeaderText = "ID_N8"
        Me.DGV.Columns(23).Width = 0
        Me.DGV.Columns(23).Visible = False

        Me.DGV.Columns(24).HeaderText = "ORDEN_N8"
        Me.DGV.Columns(24).Width = 0
        Me.DGV.Columns(24).Visible = False

        Me.DGV.Columns(25).HeaderText = "Nivel 8"
        Me.DGV.Columns(25).Width = 0
        Me.DGV.Columns(25).Visible = False
        Me.DGV.Columns(25).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(26).HeaderText = "ID_T_ES"
        Me.DGV.Columns(26).Width = 0
        Me.DGV.Columns(26).Visible = False

        Me.DGV.Columns(27).HeaderText = "D_T_ES"
        Me.DGV.Columns(27).Width = 0
        Me.DGV.Columns(27).Visible = False

        Me.DGV.Columns(28).HeaderText = "No. Orden [MIXTO]"
        Me.DGV.Columns(28).Width = 0
        Me.DGV.Columns(28).Visible = False
        Me.DGV.Columns(28).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(28).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(28).DefaultCellStyle.BackColor = Color.Khaki

        Me.DGV.Columns(29).HeaderText = "No. Orden [MILITAR]"
        Me.DGV.Columns(29).Width = 0
        Me.DGV.Columns(29).Visible = False
        Me.DGV.Columns(29).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(29).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(29).DefaultCellStyle.BackColor = Color.LightGreen

        Me.DGV.Columns(30).HeaderText = "No. Orden [PAME]"
        Me.DGV.Columns(30).Width = 50
        Me.DGV.Columns(30).Visible = True
        Me.DGV.Columns(30).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(30).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(30).DefaultCellStyle.BackColor = Color.LightBlue

        Me.DGV.Columns(31).HeaderText = "ID_CARGO_ES"
        Me.DGV.Columns(31).Width = 0
        Me.DGV.Columns(31).Visible = False

        Me.DGV.Columns(32).HeaderText = "Cargo"
        Me.DGV.Columns(32).Width = 180
        Me.DGV.Columns(32).Visible = True
        Me.DGV.Columns(32).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(33).HeaderText = "ID_GP"
        Me.DGV.Columns(33).Width = 0
        Me.DGV.Columns(33).Visible = False

        Me.DGV.Columns(34).HeaderText = "Grado Plantilla"
        Me.DGV.Columns(34).Width = 0
        Me.DGV.Columns(34).Visible = False
        Me.DGV.Columns(34).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(35).HeaderText = "ID_GR"
        Me.DGV.Columns(35).Width = 0
        Me.DGV.Columns(35).Visible = False

        Me.DGV.Columns(36).HeaderText = "Grado Real"
        Me.DGV.Columns(36).Width = 0
        Me.DGV.Columns(36).Visible = False
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

        Me.DGV.Columns(41).HeaderText = "Apellidos y Nombres"
        Me.DGV.Columns(41).Width = 200
        Me.DGV.Columns(41).Visible = True
        Me.DGV.Columns(41).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(41).DefaultCellStyle.BackColor = Color.Khaki

        Me.DGV.Columns(42).HeaderText = "Nombres y Apellidos"
        Me.DGV.Columns(42).Width = 0
        Me.DGV.Columns(42).Visible = False

        Me.DGV.Columns(43).HeaderText = "ID_SITUACION"
        Me.DGV.Columns(43).Width = 0
        Me.DGV.Columns(43).Visible = False

        Me.DGV.Columns(44).HeaderText = "Situacion"
        Me.DGV.Columns(44).Width = 0
        Me.DGV.Columns(44).Visible = False
        Me.DGV.Columns(44).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(45).HeaderText = "Jefe"
        Me.DGV.Columns(45).Width = 0
        Me.DGV.Columns(45).Visible = False
        Me.DGV.Columns(45).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(46).HeaderText = "Mixta"
        Me.DGV.Columns(46).Width = 0
        Me.DGV.Columns(46).Visible = False
        Me.DGV.Columns(46).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(47).HeaderText = "Militar"
        Me.DGV.Columns(47).Width = 0
        Me.DGV.Columns(47).Visible = False
        Me.DGV.Columns(47).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(48).HeaderText = "Pame"
        Me.DGV.Columns(48).Width = 0
        Me.DGV.Columns(48).Visible = False
        Me.DGV.Columns(48).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(49).HeaderText = "ID_ESTABLECIMIENTO"
        Me.DGV.Columns(49).Width = 0
        Me.DGV.Columns(49).Visible = False

        Me.DGV.Columns(50).HeaderText = "Establecimiento"
        Me.DGV.Columns(50).Width = 160
        Me.DGV.Columns(50).Visible = True
        Me.DGV.Columns(50).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(51).HeaderText = "D_ADMIN"
        Me.DGV.Columns(51).Width = 0
        Me.DGV.Columns(51).Visible = False
    End Sub
    Private Sub DGV_Click(sender As Object, e As EventArgs) Handles DGV.Click
        If Me.DGV.RowCount > 0 Then
            For I = 0 To Me.DGV.RowCount - 1
                If DGV.Rows(I).Selected = True Then
                    cID_M_C = Me.DGV.Item(0, I).Value
                    cORDENPAME = Me.DGV.Item(30, I).Value
                    vID_M_P = Me.DGV.Item(39, I).Value
                    cEMPLEADO = Me.DGV.Item(40, I).Value
                    bAN = Me.DGV.Item(41, I).Value
                    If Not IsDBNull(Me.DGV.Item(4, I).Value) Then : cN1 = Me.DGV.Item(4, I).Value : Else : cN1 = "-- SIN DATOS -- " : End If
                    If Not IsDBNull(Me.DGV.Item(7, I).Value) Then : cN2 = Me.DGV.Item(7, I).Value : Else : cN2 = "-- SIN DATOS -- " : End If
                    If Not IsDBNull(Me.DGV.Item(10, I).Value) Then : cN3 = Me.DGV.Item(10, I).Value : Else : cN3 = "-- SIN DATOS -- " : End If
                    If Not IsDBNull(Me.DGV.Item(13, I).Value) Then : cN4 = Me.DGV.Item(13, I).Value : Else : cN4 = "-- SIN DATOS -- " : End If
                    If Not IsDBNull(Me.DGV.Item(16, I).Value) Then : cN5 = Me.DGV.Item(16, I).Value : Else : cN5 = "-- SIN DATOS -- " : End If
                    If Not IsDBNull(Me.DGV.Item(19, I).Value) Then : cN6 = Me.DGV.Item(19, I).Value : Else : cN6 = "-- SIN DATOS -- " : End If
                    If Not IsDBNull(Me.DGV.Item(22, I).Value) Then : cN7 = Me.DGV.Item(22, I).Value : Else : cN7 = "-- SIN DATOS -- " : End If
                    If Not IsDBNull(Me.DGV.Item(25, I).Value) Then : cN8 = Me.DGV.Item(25, I).Value : Else : cN8 = "-- SIN DATOS -- " : End If
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub DGV_DoubleClick(sender As Object, e As EventArgs) Handles DGV.DoubleClick
        If vID_M_P <> 0 Then
            Me.TextBox1.Text = "0"
            Me.TextBox2.Text = "Buscar ..."
            Me.TextBox2.SelectAll()
            Me.TextBox2.Focus()
            vmcsID_M_Ct = cID_M_C
            CERRAR = False
            Me.Close()
        Else
            MsgBox("Debe seleccionar el registro que desea Visualizar", vbInformation, "Mensaje del Sistema")
        End If
    End Sub
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If Me.RadioButton1.Checked = True Then
            Me.CheckBox1.Checked = False
            Me.TextBox2.SelectAll()
            Me.TextBox2.Focus()
        End If
    End Sub
    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If Me.RadioButton2.Checked = True Then
            Me.CheckBox1.Checked = True
            Me.TextBox2.SelectAll()
            Me.TextBox2.Focus()
        End If
    End Sub
    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        If Me.RadioButton3.Checked = True Then
            Me.CheckBox1.Checked = False
            Me.TextBox2.SelectAll()
            Me.TextBox2.Focus()
        End If
    End Sub
End Class