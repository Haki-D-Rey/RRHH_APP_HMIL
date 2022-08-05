Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Aspirantes_Buscar
    Private Sub Frm_Aspirantes_Buscar_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = "0"
            Me.TextBox2.Text = "Buscar Aspirante..."
            Me.CheckBox1.Checked = True
            Me.TextBox2.SelectAll()
            Me.TextBox2.Focus()
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Aspirantes_Buscar_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = "0"
        Me.TextBox2.Text = "Buscar Aspirante..."
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
        Me.CheckBox1.Checked = True
        Me.TextBox2.SelectAll()
        Me.TextBox2.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox1.Text = "0"
        Me.TextBox2.Text = "Buscar Aspirante..."
        Me.CheckBox1.Checked = True
        Me.TextBox2.SelectAll()
        Me.TextBox2.Focus()
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
    Private Sub CARGAR_DATAGRIDVIEW()
        If Me.RadioButton1.Checked = True Then
            CADENAsql = "Select * from [VISTA SOL_MAESTRO DE SOLICITANTES] WHERE (CEDULA LIKE '%" & Me.TextBox2.Text & "%') ORDER BY CEDULA"
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
                    CADENAsql = "SELECT * FROM [dbo].[VISTA SOL_MAESTRO DE SOLICITANTES] WHERE (AN LIKE '%" & Me.TextBox2.Text & "%')"
                ElseIf tamañoArreglo = 1 Then
                    CADENAsql = "SELECT * FROM [dbo].[VISTA SOL_MAESTRO DE SOLICITANTES] WHERE AN LIKE '%" & buscar & "%'"
                ElseIf tamañoArreglo > 1 Then
                    CADENAsql = "SELECT * FROM [dbo].[VISTA SOL_MAESTRO DE SOLICITANTES] WHERE AN LIKE '%" & ArrCadena(0) & "%' "
                    For index = 0 To ArrCadena.GetUpperBound(0)
                        CADENAsql = CADENAsql & " UNION SELECT * FROM [dbo].[VISTA SOL_MAESTRO DE SOLICITANTES] WHERE AN LIKE '%" & ArrCadena(index) & "%'"
                    Next
                End If
                CADENAsql = CADENAsql & "  ORDER BY AN"
                GoTo SALTO
            End If
            If Me.CheckBox1.Checked = True Then
                CADENAsql = "Select * from [VISTA SOL_MAESTRO DE SOLICITANTES] WHERE (AN LIKE '%" & Me.TextBox2.Text & "%') ORDER BY AN"
            End If
        End If
SALTO:
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA SOL_MAESTRO DE SOLICITANTES]")
        DGV.DataSource = MiDataSet.Tables("[VISTA SOL_MAESTRO DE SOLICITANTES]")
        Me.Label2.Text = "Total de Registros: " & DGV.RowCount
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

        Me.DGV.Columns(1).HeaderText = "Cedula"
        Me.DGV.Columns(1).Width = 100
        Me.DGV.Columns(1).Visible = True
        Me.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(1).DefaultCellStyle.BackColor = Color.LightBlue

        Me.DGV.Columns(2).HeaderText = "APELLIDOS"
        Me.DGV.Columns(2).Width = 0
        Me.DGV.Columns(2).Visible = False

        Me.DGV.Columns(3).HeaderText = "NOMBRES"
        Me.DGV.Columns(3).Width = 0
        Me.DGV.Columns(3).Visible = False

        Me.DGV.Columns(4).HeaderText = "Aspirante"
        Me.DGV.Columns(4).Width = 230
        Me.DGV.Columns(4).Visible = True
        Me.DGV.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(5).HeaderText = "Fecha Nacimiento"
        Me.DGV.Columns(5).Width = 100
        Me.DGV.Columns(5).Visible = True
        Me.DGV.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(5).DefaultCellStyle.BackColor = Color.LightGreen

        Me.DGV.Columns(6).HeaderText = ""
        Me.DGV.Columns(6).Width = 0
        Me.DGV.Columns(6).Visible = False
        Me.DGV.Columns(6).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(7).HeaderText = "Perfil"
        Me.DGV.Columns(7).Width = 0
        Me.DGV.Columns(7).Visible = False
        Me.DGV.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(8).HeaderText = "ID_T_SEXO"
        Me.DGV.Columns(8).Width = 0
        Me.DGV.Columns(8).Visible = False

        Me.DGV.Columns(9).HeaderText = "Sexo"
        Me.DGV.Columns(9).Width = 90
        Me.DGV.Columns(9).Visible = True
        Me.DGV.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(10).HeaderText = "Direccion 1"
        Me.DGV.Columns(10).Width = 0
        Me.DGV.Columns(10).Visible = False
        Me.DGV.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(11).HeaderText = "Direccion 2"
        Me.DGV.Columns(11).Width = 0
        Me.DGV.Columns(11).Visible = False
        Me.DGV.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(12).HeaderText = "Telefono Claro"
        Me.DGV.Columns(12).Width = 80
        Me.DGV.Columns(12).Visible = True
        Me.DGV.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(13).HeaderText = "Telefono Tigo"
        Me.DGV.Columns(13).Width = 0
        Me.DGV.Columns(13).Visible = False
        Me.DGV.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(14).HeaderText = "Correo 1"
        Me.DGV.Columns(14).Width = 0
        Me.DGV.Columns(14).Visible = FALSE
        Me.DGV.Columns(14).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(15).HeaderText = "Correo 2"
        Me.DGV.Columns(15).Width = 0
        Me.DGV.Columns(15).Visible = False
        Me.DGV.Columns(15).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(16).HeaderText = "Whatsapp"
        Me.DGV.Columns(16).Width = 0
        Me.DGV.Columns(16).Visible = False
        Me.DGV.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(16).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(17).HeaderText = "Padre"
        Me.DGV.Columns(17).Width = 0
        Me.DGV.Columns(17).Visible = False
        Me.DGV.Columns(17).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(18).HeaderText = "Madre"
        Me.DGV.Columns(18).Width = 0
        Me.DGV.Columns(18).Visible = False
        Me.DGV.Columns(18).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(19).HeaderText = "Referencia Academica"
        Me.DGV.Columns(19).Width = 0
        Me.DGV.Columns(19).Visible = False
        Me.DGV.Columns(19).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(20).HeaderText = "Referencia Estudios"
        Me.DGV.Columns(20).Width = 0
        Me.DGV.Columns(20).Visible = False
        Me.DGV.Columns(20).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(21).HeaderText = "Recomendado por"
        Me.DGV.Columns(21).Width = 0
        Me.DGV.Columns(21).Visible = False
        Me.DGV.Columns(21).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(22).HeaderText = "Observaciones"
        Me.DGV.Columns(22).Width = 0
        Me.DGV.Columns(22).Visible = False
        Me.DGV.Columns(22).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(23).HeaderText = "ID_ESTADO_R"
        Me.DGV.Columns(23).Width = 0
        Me.DGV.Columns(23).Visible = False

        Me.DGV.Columns(24).HeaderText = "Fecha Ingreso"
        Me.DGV.Columns(24).Width = 70
        Me.DGV.Columns(24).Visible = True
        Me.DGV.Columns(24).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(24).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(24).DefaultCellStyle.BackColor = Color.Khaki
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
    Private Sub DGV_Click(sender As Object, e As EventArgs) Handles DGV.Click
        If Me.DGV.RowCount > 0 Then
            For I = 0 To Me.DGV.RowCount - 1
                If DGV.Rows(I).Selected = True Then
                    aspID_S = Me.DGV.Item(0, I).Value
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub DGV_DoubleClick(sender As Object, e As EventArgs) Handles DGV.DoubleClick
        If aspID_S <> 0 Then
            Me.TextBox1.Text = "0"
            Me.TextBox2.Text = "Buscar Aspirante..."
            Me.TextBox2.SelectAll()
            Me.TextBox2.Focus()
            CERRAR = False
            Me.Close()
        End If
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Me.TextBox2.SelectAll()
        Me.TextBox2.Focus()
    End Sub
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        Me.TextBox2.SelectAll()
        Me.TextBox2.Focus()
    End Sub
    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        Me.TextBox2.SelectAll()
        Me.TextBox2.Focus()
    End Sub
End Class