Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Movimientos_Exp_Add_Buscar_Norden
    Private Sub Frm_Movimientos_Exp_Add_Buscar_Norden_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox3.SelectAll()
            Me.TextBox3.Focus()
            maOM = ""
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Movimientos_Exp_Add_Buscar_Norden_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = "0"
        Me.TextBox3.Text = "Buscar ..."
        Me.RadioButton2.Checked = True
        Me.CheckBox1.Checked = True
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
        Me.TextBox3.SelectAll()
        Me.TextBox3.Focus()
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Me.TextBox3.SelectAll()
        Me.TextBox3.Focus()
        maOM = ""
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW()
        If Me.RadioButton2.Checked = True Then
            If Me.CheckBox1.Checked = False Then
                Dim buscar As String = ""
                Dim tamañoArreglo As Integer = 0
                buscar = Me.TextBox3.Text.Trim
                Dim ArrCadena As String() = buscar.Split(" ")
                tamañoArreglo = ArrCadena.GetUpperBound(0) + 1
                If tamañoArreglo = 0 Then
                    CADENAsql = "Select ID_M_C, N_ORDEN_MIXTA, D_CARGO_ES, D_SITUACION, AN FROM [dbo].[VISTA MAESTRO DE CARGOS SIT] WHERE (D_CARGO_ES LIKE '%" & Me.TextBox3.Text & "%')"
                ElseIf tamañoArreglo = 1 Then
                    CADENAsql = "Select ID_M_C, N_ORDEN_MIXTA, D_CARGO_ES, D_SITUACION, AN FROM [dbo].[VISTA MAESTRO DE CARGOS SIT] WHERE (D_CARGO_ES LIKE '%" & buscar & "%')"
                ElseIf tamañoArreglo > 1 Then
                    CADENAsql = "Select ID_M_C, N_ORDEN_MIXTA, D_CARGO_ES, D_SITUACION, AN FROM [dbo].[VISTA MAESTRO DE CARGOS SIT] WHERE (D_CARGO_ES LIKE '%" & ArrCadena(0) & "%')"
                    For index = 0 To ArrCadena.GetUpperBound(0)
                        CADENAsql = CADENAsql & " UNION Select ID_M_C, N_ORDEN_MIXTA, D_CARGO_ES, AND D_SITUACION FROM [dbo].[VISTA MAESTRO DE CARGOS SIT] WHERE (D_CARGO_ES LIKE '%" & ArrCadena(index) & "%')"
                    Next
                End If
                CADENAsql = CADENAsql & "  ORDER BY N_ORDEN_MIXTA"
                GoTo SALTO
            End If
            If Me.CheckBox1.Checked = True Then
                CADENAsql = "Select ID_M_C, N_ORDEN_MIXTA, D_CARGO_ES, D_SITUACION, AN from [VISTA MAESTRO DE CARGOS SIT] WHERE (D_CARGO_ES LIKE '%" & Me.TextBox3.Text & "%') ORDER BY N_ORDEN_MIXTA"
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

        Me.DGV.Columns(1).HeaderText = "Orden Mixto"
        Me.DGV.Columns(1).Width = 40
        Me.DGV.Columns(1).Visible = True
        Me.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(1).DefaultCellStyle.BackColor = Color.Khaki

        Me.DGV.Columns(2).HeaderText = "Cargo"
        Me.DGV.Columns(2).Width = 200
        Me.DGV.Columns(2).Visible = True
        Me.DGV.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(3).HeaderText = "Situacion"
        Me.DGV.Columns(3).Width = 60
        Me.DGV.Columns(3).Visible = True
        Me.DGV.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(4).HeaderText = "Apellidos y Nombres"
        Me.DGV.Columns(4).Width = 200
        Me.DGV.Columns(4).Visible = True
        Me.DGV.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
    End Sub
    Private Sub DGV_RowPrePaint(sender As Object, e As DataGridViewRowPrePaintEventArgs) Handles DGV.RowPrePaint
        Call FORMATO_FILAS_DATAGRIDVIEW()
    End Sub
    Private Sub FORMATO_FILAS_DATAGRIDVIEW()
        Dim style As New DataGridViewCellStyle
        style.Font = New Font(DGV.Font, FontStyle.Bold)
        For Each Row As DataGridViewRow In DGV.Rows
            If Row.Cells(3).Value = "VACANTE" Then
                Row.DefaultCellStyle.ForeColor = Color.Blue
            End If
        Next
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If Len(Me.TextBox3.Text) = 0 Then
            MsgBox("Debe digitar correctamente el Texto que desea Buscar", vbInformation, "Mensaje del sistema")
            Me.TextBox3.Focus()
        Else
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Me.TextBox3.SelectAll()
            Me.TextBox3.Focus()
        End If
    End Sub
    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If (e.KeyCode = Keys.Enter) Then
            If Len(Me.TextBox3.Text) = 0 Then
                MsgBox("Debe digitar correctamente el Texto que desea Buscar", vbInformation, "Mensaje del sistema")
                Me.TextBox3.Focus()
            Else
                Call CARGAR_DATAGRIDVIEW()
                Call ARMAR_DATAGRIDVIEW()
                Me.TextBox3.SelectAll()
                Me.TextBox3.Focus()
            End If
        End If
    End Sub
    Private Sub DGV_Click(sender As Object, e As EventArgs) Handles DGV.Click
        If Me.DGV.RowCount > 0 Then
            For I = 0 To Me.DGV.RowCount - 1
                If DGV.Rows(I).Selected = True Then
                    idmaOM = Me.DGV.Item(0, I).Value
                    maOM = Me.DGV.Item(1, I).Value
                    Exit Sub
                Else
                    idmaOM = 0
                    maOM = ""
                End If
            Next
        End If
    End Sub
    Private Sub DGV_DoubleClick(sender As Object, e As EventArgs) Handles DGV.DoubleClick
        If maOM <> "" Then
            Me.TextBox1.Text = "0"
            Me.TextBox3.Text = "Buscar ..."
            Me.TextBox3.SelectAll()
            Me.TextBox3.Focus()
            CERRAR = False
            Me.Close()
        Else
            MsgBox("Debe seleccionar el registro que desea Visualizar", vbInformation, "Mensaje del Sistema")
        End If
    End Sub
    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If Me.RadioButton2.Checked = True Then
            Me.CheckBox1.Checked = True
            Me.TextBox3.SelectAll()
            Me.TextBox3.Focus()
        End If
    End Sub
End Class