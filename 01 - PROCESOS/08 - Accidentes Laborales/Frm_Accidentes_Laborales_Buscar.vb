Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Accidentes_Laborales_Buscar
    Private Sub Frm_Accidentes_Laborales_Buscar_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox2.SelectAll()
            Me.TextBox2.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Accidentes_Laborales_Buscar_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = "0"
        Me.TextBox2.Text = "Buscar Empleado..."
        Me.RadioButton1.Checked = True
        Me.RadioButton2.Checked = False
        Me.CheckBox1.Checked = False
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
        Me.TextBox2.SelectAll()
        Me.TextBox2.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
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
        CADENAsql1 = "ID_M_P, FEC_NAC, FEC_ING_PAME, FEC_ING_EN, CODIGO, APELLIDOS, NOMBRES, AN, ID_T_SEXO, D_T_SEXO, " &
         "DIRECCION_1, DIRECCION_2, NSS, N_CEDULA, N_LIC, N_MINSA, LIC_IMAGEN, TELEFONO_1, TELEFONO_2, CONVENCIONAL, " &
         "CORREO_1, CORREO_2, WHATSAPP, FACEBOOK, TWITTER, PESO, ESTATURA, ID_TEZ, D_TEZ, ID_CABELLO, " &
         "D_CABELLO, ID_N_ACADEMICO, D_N_ACADEMICO, ID_N_PROFESIONAL, D_N_PROFESIONAL, ID_NACIONALIDAD_N, D_NACIONALIDAD_N, ID_DPTO_N, D_DPTO_N, ID_M_N, " &
         "D_M_N, LUGAR_N, ID_NACIONALIDAD_D, D_NACIONALIDAD_D, ID_DPTO_D, D_DPTO_D, ID_M_D, D_M_D, LUGAR_D, ID_E_CIVIL, D_E_CIVIL, " &
         "ID_T_HORARIO, D_T_HORARIO , N_C_BANCO, EMPLEADO_ANT, ID_T_AFILIACION, D_T_AFILIACION, ID_T_PLAZA, D_T_PLAZA, D_N1, D_N2, " &
         "ID_ESTADO_P, D_ESTADO_P, BECADO, EDAD, ANTIGUEDAD_PAME, ANTIGUEDAD_EN"
        If Me.RadioButton1.Checked = True Then
            CADENAsql = "Select " & CADENAsql1 & " from [VISTA MAESTRO DE PERSONAS HIGIENE] WHERE (CODIGO LIKE '%" & Me.TextBox2.Text & "%') ORDER BY AN"
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
                    CADENAsql = "Select " & CADENAsql1 & " FROM [dbo].[VISTA MAESTRO DE PERSONAS HIGIENE] WHERE (CODIGO <> '0000000000') AND (AN LIKE '%" & Me.TextBox2.Text & "%')"
                ElseIf tamañoArreglo = 1 Then
                    CADENAsql = "Select " & CADENAsql1 & " FROM [dbo].[VISTA MAESTRO DE PERSONAS HIGIENE] WHERE (CODIGO <> '0000000000') AND AN LIKE '%" & buscar & "%'"
                ElseIf tamañoArreglo > 1 Then
                    CADENAsql = "Select " & CADENAsql1 & " FROM [dbo].[VISTA MAESTRO DE PERSONAS HIGIENE] WHERE (CODIGO <> '0000000000') AND AN LIKE '%" & ArrCadena(0) & "%' "
                    For index = 0 To ArrCadena.GetUpperBound(0)
                        CADENAsql = CADENAsql & " UNION Select " & CADENAsql1 & " FROM [dbo].[VISTA MAESTRO DE PERSONAS HIGIENE] WHERE (CODIGO <> '0000000000') AND AN LIKE '%" & ArrCadena(index) & "%'"
                    Next
                End If
                CADENAsql = CADENAsql & "  ORDER BY AN"
                GoTo SALTO
            End If
            If Me.CheckBox1.Checked = True Then
                CADENAsql = "Select " & CADENAsql1 & " from [VISTA MAESTRO DE PERSONAS HIGIENE] WHERE (AN LIKE '%" & Me.TextBox2.Text & "%') ORDER BY AN"
            End If
        End If
SALTO:
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE PERSONAS HIGIENE]")
        DGV.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE PERSONAS HIGIENE]")
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

        Me.DGV.Columns(1).HeaderText = "Fecha Nacimiento"
        Me.DGV.Columns(1).Width = 80
        Me.DGV.Columns(1).Visible = True
        Me.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(2).HeaderText = "Fecha Ing. PAME"
        Me.DGV.Columns(2).Width = 80
        Me.DGV.Columns(2).Visible = True
        Me.DGV.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(3).HeaderText = "FEC_ING_EN"
        Me.DGV.Columns(3).Width = 0
        Me.DGV.Columns(3).Visible = False

        Me.DGV.Columns(4).HeaderText = "Codigo"
        Me.DGV.Columns(4).Width = 70
        Me.DGV.Columns(4).Visible = True
        Me.DGV.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(5).HeaderText = "APELLIDOS"
        Me.DGV.Columns(5).Width = 0
        Me.DGV.Columns(5).Visible = False

        Me.DGV.Columns(6).HeaderText = "NOMBRES"
        Me.DGV.Columns(6).Width = 0
        Me.DGV.Columns(6).Visible = False

        Me.DGV.Columns(7).HeaderText = "Apellidos y Nombres"
        Me.DGV.Columns(7).Width = 150
        Me.DGV.Columns(7).Visible = True
        Me.DGV.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(8).HeaderText = ""
        Me.DGV.Columns(8).Width = 0
        Me.DGV.Columns(8).Visible = False

        Me.DGV.Columns(9).HeaderText = ""
        Me.DGV.Columns(9).Width = 0
        Me.DGV.Columns(9).Visible = False

        Me.DGV.Columns(10).HeaderText = ""
        Me.DGV.Columns(10).Width = 0
        Me.DGV.Columns(10).Visible = False

        Me.DGV.Columns(11).HeaderText = ""
        Me.DGV.Columns(11).Width = 0
        Me.DGV.Columns(11).Visible = False

        Me.DGV.Columns(12).HeaderText = "Nss"
        Me.DGV.Columns(12).Width = 70
        Me.DGV.Columns(12).Visible = True
        Me.DGV.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(13).HeaderText = "Cédula"
        Me.DGV.Columns(13).Width = 80
        Me.DGV.Columns(13).Visible = True
        Me.DGV.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(14).HeaderText = "0"
        Me.DGV.Columns(14).Width = 0
        Me.DGV.Columns(14).Visible = False
        Me.DGV.Columns(14).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(15).HeaderText = ""
        Me.DGV.Columns(15).Width = 0
        Me.DGV.Columns(15).Visible = False

        Me.DGV.Columns(16).HeaderText = ""
        Me.DGV.Columns(16).Width = 0
        Me.DGV.Columns(16).Visible = False

        Me.DGV.Columns(17).HeaderText = ""
        Me.DGV.Columns(17).Width = 0
        Me.DGV.Columns(17).Visible = False

        Me.DGV.Columns(18).HeaderText = ""
        Me.DGV.Columns(18).Width = 0
        Me.DGV.Columns(18).Visible = False

        Me.DGV.Columns(19).HeaderText = ""
        Me.DGV.Columns(19).Width = 0
        Me.DGV.Columns(19).Visible = False

        Me.DGV.Columns(20).HeaderText = ""
        Me.DGV.Columns(20).Width = 0
        Me.DGV.Columns(20).Visible = False

        Me.DGV.Columns(21).HeaderText = ""
        Me.DGV.Columns(21).Width = 0
        Me.DGV.Columns(21).Visible = False

        Me.DGV.Columns(22).HeaderText = ""
        Me.DGV.Columns(22).Width = 0
        Me.DGV.Columns(22).Visible = False

        Me.DGV.Columns(23).HeaderText = ""
        Me.DGV.Columns(23).Width = 0
        Me.DGV.Columns(23).Visible = False

        Me.DGV.Columns(24).HeaderText = ""
        Me.DGV.Columns(24).Width = 0
        Me.DGV.Columns(24).Visible = False

        Me.DGV.Columns(25).HeaderText = ""
        Me.DGV.Columns(25).Width = 0
        Me.DGV.Columns(25).Visible = False

        Me.DGV.Columns(26).HeaderText = ""
        Me.DGV.Columns(26).Width = 0
        Me.DGV.Columns(26).Visible = False

        Me.DGV.Columns(27).HeaderText = ""
        Me.DGV.Columns(27).Width = 0
        Me.DGV.Columns(27).Visible = False

        Me.DGV.Columns(28).HeaderText = ""
        Me.DGV.Columns(28).Width = 0
        Me.DGV.Columns(28).Visible = False

        Me.DGV.Columns(29).HeaderText = ""
        Me.DGV.Columns(29).Width = 0
        Me.DGV.Columns(29).Visible = False

        Me.DGV.Columns(30).HeaderText = ""
        Me.DGV.Columns(30).Width = 0
        Me.DGV.Columns(30).Visible = False

        Me.DGV.Columns(31).HeaderText = ""
        Me.DGV.Columns(31).Width = 0
        Me.DGV.Columns(31).Visible = False

        Me.DGV.Columns(32).HeaderText = ""
        Me.DGV.Columns(32).Width = 0
        Me.DGV.Columns(32).Visible = False

        Me.DGV.Columns(33).HeaderText = ""
        Me.DGV.Columns(33).Width = 0
        Me.DGV.Columns(33).Visible = False

        Me.DGV.Columns(34).HeaderText = ""
        Me.DGV.Columns(34).Width = 0
        Me.DGV.Columns(34).Visible = False

        Me.DGV.Columns(35).HeaderText = ""
        Me.DGV.Columns(35).Width = 0
        Me.DGV.Columns(35).Visible = False

        Me.DGV.Columns(36).HeaderText = ""
        Me.DGV.Columns(36).Width = 0
        Me.DGV.Columns(36).Visible = False

        Me.DGV.Columns(37).HeaderText = ""
        Me.DGV.Columns(37).Width = 0
        Me.DGV.Columns(37).Visible = False

        Me.DGV.Columns(38).HeaderText = ""
        Me.DGV.Columns(38).Width = 0
        Me.DGV.Columns(38).Visible = False

        Me.DGV.Columns(39).HeaderText = ""
        Me.DGV.Columns(39).Width = 0
        Me.DGV.Columns(39).Visible = False

        Me.DGV.Columns(40).HeaderText = ""
        Me.DGV.Columns(40).Width = 0
        Me.DGV.Columns(40).Visible = False

        Me.DGV.Columns(41).HeaderText = ""
        Me.DGV.Columns(41).Width = 0
        Me.DGV.Columns(41).Visible = False

        Me.DGV.Columns(42).HeaderText = ""
        Me.DGV.Columns(42).Width = 0
        Me.DGV.Columns(42).Visible = False

        Me.DGV.Columns(43).HeaderText = ""
        Me.DGV.Columns(43).Width = 0
        Me.DGV.Columns(43).Visible = False

        Me.DGV.Columns(44).HeaderText = ""
        Me.DGV.Columns(44).Width = 0
        Me.DGV.Columns(44).Visible = False

        Me.DGV.Columns(45).HeaderText = ""
        Me.DGV.Columns(45).Width = 0
        Me.DGV.Columns(45).Visible = False

        Me.DGV.Columns(46).HeaderText = ""
        Me.DGV.Columns(46).Width = 0
        Me.DGV.Columns(46).Visible = False

        Me.DGV.Columns(47).HeaderText = ""
        Me.DGV.Columns(47).Width = 0
        Me.DGV.Columns(47).Visible = False

        Me.DGV.Columns(48).HeaderText = ""
        Me.DGV.Columns(48).Width = 0
        Me.DGV.Columns(48).Visible = False

        Me.DGV.Columns(49).HeaderText = ""
        Me.DGV.Columns(49).Width = 0
        Me.DGV.Columns(49).Visible = False

        Me.DGV.Columns(50).HeaderText = ""
        Me.DGV.Columns(50).Width = 0
        Me.DGV.Columns(50).Visible = False

        Me.DGV.Columns(51).HeaderText = ""
        Me.DGV.Columns(51).Width = 0
        Me.DGV.Columns(51).Visible = False

        Me.DGV.Columns(52).HeaderText = ""
        Me.DGV.Columns(52).Width = 0
        Me.DGV.Columns(52).Visible = False

        Me.DGV.Columns(53).HeaderText = ""
        Me.DGV.Columns(53).Width = 0
        Me.DGV.Columns(53).Visible = False

        Me.DGV.Columns(54).HeaderText = ""
        Me.DGV.Columns(54).Width = 0
        Me.DGV.Columns(54).Visible = False

        Me.DGV.Columns(55).HeaderText = ""
        Me.DGV.Columns(55).Width = 0
        Me.DGV.Columns(55).Visible = False

        Me.DGV.Columns(56).HeaderText = ""
        Me.DGV.Columns(56).Width = 0
        Me.DGV.Columns(56).Visible = False

        Me.DGV.Columns(57).HeaderText = ""
        Me.DGV.Columns(57).Width = 0
        Me.DGV.Columns(57).Visible = False

        Me.DGV.Columns(58).HeaderText = ""
        Me.DGV.Columns(58).Width = 0
        Me.DGV.Columns(58).Visible = False

        Me.DGV.Columns(59).HeaderText = ""
        Me.DGV.Columns(59).Width = 0
        Me.DGV.Columns(59).Visible = False

        Me.DGV.Columns(60).HeaderText = ""
        Me.DGV.Columns(60).Width = 0
        Me.DGV.Columns(60).Visible = False

        Me.DGV.Columns(61).HeaderText = ""
        Me.DGV.Columns(61).Width = 0
        Me.DGV.Columns(61).Visible = False

        Me.DGV.Columns(62).HeaderText = "Estado"
        Me.DGV.Columns(62).Width = 70
        Me.DGV.Columns(62).Visible = True
        Me.DGV.Columns(62).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(63).HeaderText = ""
        Me.DGV.Columns(63).Width = 0
        Me.DGV.Columns(63).Visible = False

        Me.DGV.Columns(64).HeaderText = "Edad"
        Me.DGV.Columns(64).Width = 50
        Me.DGV.Columns(64).Visible = True
        Me.DGV.Columns(64).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(64).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(65).HeaderText = ""
        Me.DGV.Columns(65).Width = 0
        Me.DGV.Columns(65).Visible = False

        Me.DGV.Columns(66).HeaderText = ""
        Me.DGV.Columns(66).Width = 0
        Me.DGV.Columns(66).Visible = False
    End Sub
    Private Sub DGV_Click(sender As Object, e As EventArgs) Handles DGV.Click
        If Me.DGV.RowCount > 0 Then
            For I = 0 To Me.DGV.RowCount - 1
                If DGV.Rows(I).Selected = True Then
                    h1ID_M_P = Me.DGV.Item(0, I).Value
                    h1CODIGO = Me.DGV.Item(4, I).Value
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub DGV_DoubleClick(sender As Object, e As EventArgs) Handles DGV.DoubleClick
        If h1ID_M_P <> 0 Then
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
    Private Sub DGV_RowPrePaint(sender As Object, e As DataGridViewRowPrePaintEventArgs) Handles DGV.RowPrePaint
        Call FORMATO_FILAS_DATAGRIDVIEW()
    End Sub
    Private Sub FORMATO_FILAS_DATAGRIDVIEW()
        Dim style As New DataGridViewCellStyle
        style.Font = New Font(DGV.Font, FontStyle.Bold)
        For Each Row As DataGridViewRow In DGV.Rows
            If Row.Cells(62).Value = "NO ACTIVO" Then
                Row.DefaultCellStyle.ForeColor = Color.Red
            End If
        Next
    End Sub
End Class