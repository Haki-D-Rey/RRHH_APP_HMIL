Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Movimientos_Exp_Add_Seleccionar_Recurso_Reingreso
    Private Sub Frm_Movimientos_Exp_Add_Seleccionar_Recurso_Reingreso_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = "0"
        Me.TextBox2.Text = "Buscar..."
        Me.RadioButton1.Checked = False
        Me.RadioButton2.Checked = True
        Me.CheckBox1.Checked = True
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
        Me.TextBox2.SelectAll()
        Me.TextBox2.Focus()
    End Sub
    Private Sub Frm_Movimientos_Exp_Add_Seleccionar_Recurso_Reingreso_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox2.SelectAll()
            Me.TextBox2.Focus()
            CERRAR = True
            Me.Close()
        End If
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
                CADENAsql = "SELECT FEC_ING_PAME, ID_M_P, CODIGO, AN, D_T_SEXO, NSS, N_CEDULA FROM [dbo].[VISTA MAESTRO DE PERSONAS] WHERE (CODIGO LIKE '%" & Me.TextBox2.Text & "%') AND (ID_ESTADO_P = 2) AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY AN"
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
                        CADENAsql = "SELECT FEC_ING_PAME, ID_M_P, CODIGO, AN, D_T_SEXO, NSS, N_CEDULA  FROM [dbo].[VISTA MAESTRO DE PERSONAS] WHERE (AN LIKE '%" & Me.TextBox2.Text & "%') AND D_ADMIN = '" & isuUSUARIO & "'"
                    ElseIf tamañoArreglo = 1 Then
                        CADENAsql = "SELECT FEC_ING_PAME, ID_M_P, CODIGO, AN, D_T_SEXO, NSS, N_CEDULA  FROM [dbo].[VISTA MAESTRO DE PERSONAS] WHERE AN LIKE '%" & buscar & "%' AND D_ADMIN = '" & isuUSUARIO & "'"
                    ElseIf tamañoArreglo > 1 Then
                        CADENAsql = "SELECT FEC_ING_PAME, ID_M_P, CODIGO, AN, D_T_SEXO, NSS, N_CEDULA  FROM [dbo].[VISTA MAESTRO DE PERSONAS] WHERE AN LIKE '%" & ArrCadena(0) & "%' AND D_ADMIN = '" & isuUSUARIO & "'"
                        For index = 0 To ArrCadena.GetUpperBound(0)
                            CADENAsql = CADENAsql & " UNION SELECT FEC_ING_PAME, ID_M_P, CODIGO, AN, D_T_SEXO, NSS, N_CEDULA  FROM [dbo].[VISTA MAESTRO DE PERSONAS] WHERE AN LIKE '%" & ArrCadena(index) & "%' AND D_ADMIN = '" & isuUSUARIO & "'"
                        Next
                    End If
                    CADENAsql = CADENAsql & "  ORDER BY AN"
                    GoTo SALTO
                End If
                If Me.CheckBox1.Checked = True Then
                    CADENAsql = "SELECT FEC_ING_PAME, ID_M_P, CODIGO, AN, D_T_SEXO, NSS, N_CEDULA  FROM [dbo].[VISTA MAESTRO DE PERSONAS] WHERE (AN LIKE '%" & Me.TextBox2.Text & "%') AND (ID_ESTADO_P = 2) AND D_ADMIN = '" & isuUSUARIO & "'ORDER BY AN"
                End If
            End If
        Else
            If Me.RadioButton1.Checked = True Then
                CADENAsql = "SELECT FEC_ING_PAME, ID_M_P, CODIGO, AN, D_T_SEXO, NSS, N_CEDULA FROM [dbo].[VISTA MAESTRO DE PERSONAS] WHERE (CODIGO LIKE '%" & Me.TextBox2.Text & "%') AND (ID_ESTADO_P = 2) ORDER BY AN"
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
                        CADENAsql = "SELECT FEC_ING_PAME, ID_M_P, CODIGO, AN, D_T_SEXO, NSS, N_CEDULA  FROM [dbo].[VISTA MAESTRO DE PERSONAS] WHERE (AN LIKE '%" & Me.TextBox2.Text & "%')"
                    ElseIf tamañoArreglo = 1 Then
                        CADENAsql = "SELECT FEC_ING_PAME, ID_M_P, CODIGO, AN, D_T_SEXO, NSS, N_CEDULA  FROM [dbo].[VISTA MAESTRO DE PERSONAS] WHERE AN LIKE '%" & buscar & "%'"
                    ElseIf tamañoArreglo > 1 Then
                        CADENAsql = "SELECT FEC_ING_PAME, ID_M_P, CODIGO, AN, D_T_SEXO, NSS, N_CEDULA  FROM [dbo].[VISTA MAESTRO DE PERSONAS] WHERE AN LIKE '%" & ArrCadena(0) & "%'"
                        For index = 0 To ArrCadena.GetUpperBound(0)
                            CADENAsql = CADENAsql & " UNION SELECT FEC_ING_PAME, ID_M_P, CODIGO, AN, D_T_SEXO, NSS, N_CEDULA  FROM [dbo].[VISTA MAESTRO DE PERSONAS] WHERE AN LIKE '%" & ArrCadena(index) & "%'"
                        Next
                    End If
                    CADENAsql = CADENAsql & "  ORDER BY AN"
                    GoTo SALTO
                End If
                If Me.CheckBox1.Checked = True Then
                    CADENAsql = "SELECT FEC_ING_PAME, ID_M_P, CODIGO, AN, D_T_SEXO, NSS, N_CEDULA  FROM [dbo].[VISTA MAESTRO DE PERSONAS] WHERE (AN LIKE '%" & Me.TextBox2.Text & "%') AND (ID_ESTADO_P = 2) ORDER BY AN"
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
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE PERSONAS]")
        DGV.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE PERSONAS]")
        Me.Label2.Text = DGV.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW()
        Me.DGV.Columns(0).HeaderText = "Fec. Ing. PAME"
        Me.DGV.Columns(0).Width = 100
        Me.DGV.Columns(0).Visible = True
        Me.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(1).HeaderText = "ID_M_P"
        Me.DGV.Columns(1).Width = 0
        Me.DGV.Columns(1).Visible = False

        Me.DGV.Columns(2).HeaderText = "Codigo"
        Me.DGV.Columns(2).Width = 80
        Me.DGV.Columns(2).Visible = True
        Me.DGV.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(3).HeaderText = "Apellidos y Nombres"
        Me.DGV.Columns(3).Width = 220
        Me.DGV.Columns(3).Visible = True
        Me.DGV.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(4).HeaderText = "Sexo"
        Me.DGV.Columns(4).Width = 100
        Me.DGV.Columns(4).Visible = True
        Me.DGV.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(5).HeaderText = "Nss"
        Me.DGV.Columns(5).Width = 80
        Me.DGV.Columns(5).Visible = True
        Me.DGV.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(6).HeaderText = "Cedula"
        Me.DGV.Columns(6).Width = 100
        Me.DGV.Columns(6).Visible = True
        Me.DGV.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

    End Sub
    Private Sub DGV_Click(sender As Object, e As EventArgs) Handles DGV.Click
        If Me.DGV.RowCount > 0 Then
            For I = 0 To Me.DGV.RowCount - 1
                If DGV.Rows(I).Selected = True Then
                    'cID_M_C = Me.DGV.Item(0, I).Value
                    'cORDENPAME = Me.DGV.Item(21, I).Value
                    cFECHAING = Mid(Me.DGV.Item(0, I).Value, 1, 10)
                    vID_M_P = Me.DGV.Item(1, I).Value
                    cEMPLEADO = Me.DGV.Item(2, I).Value
                    bAN = Me.DGV.Item(3, I).Value
                    'If Not IsDBNull(Me.DGV.Item(4, I).Value) Then : cN1 = Me.DGV.Item(4, I).Value : Else : cN1 = "-- SIN DATOS -- " : End If
                    'If Not IsDBNull(Me.DGV.Item(7, I).Value) Then : cN2 = Me.DGV.Item(7, I).Value : Else : cN2 = "-- SIN DATOS -- " : End If
                    'If Not IsDBNull(Me.DGV.Item(10, I).Value) Then : cN3 = Me.DGV.Item(10, I).Value : Else : cN3 = "-- SIN DATOS -- " : End If
                    'If Not IsDBNull(Me.DGV.Item(13, I).Value) Then : cN4 = Me.DGV.Item(13, I).Value : Else : cN4 = "-- SIN DATOS -- " : End If
                    'If Not IsDBNull(Me.DGV.Item(16, I).Value) Then : cN5 = Me.DGV.Item(16, I).Value : Else : cN5 = "-- SIN DATOS -- " : End If
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
            'vmcsID_M_Ct = cID_M_C
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
End Class