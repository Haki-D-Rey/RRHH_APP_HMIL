Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Aspirantes_Cargos
    Private Sub Frm_Aspirantes_Cargos_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox3.SelectAll()
            Me.TextBox3.Focus()
            maCARGO = ""
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Aspirantes_Cargos_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = "0"
        Me.TextBox3.Text = "Buscar ..."
        Me.RadioButton2.Checked = True
        Me.CheckBox1.Checked = True
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
        Me.TextBox3.SelectAll()
        Me.TextBox3.Focus()
    End Sub

    Private Sub Button7_Click(sender As Object, e As System.EventArgs) Handles Button7.Click
        Me.TextBox3.SelectAll()
        Me.TextBox3.Focus()
        maCARGO = ""
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
                    CADENAsql = "Select DISTINCT DESCRIPCION FROM [dbo].[CAT DE CARGOS DE ESTRUCTURA] WHERE (DESCRIPCION LIKE '%" & Me.TextBox3.Text & "%')"
                ElseIf tamañoArreglo = 1 Then
                    CADENAsql = "Select DISTINCT DESCRIPCION FROM [dbo].[CAT DE CARGOS DE ESTRUCTURA] WHERE (DESCRIPCION LIKE '%" & buscar & "%')"
                ElseIf tamañoArreglo > 1 Then
                    CADENAsql = "Select DISTINCT DESCRIPCION FROM [dbo].[CAT DE CARGOS DE ESTRUCTURA] WHERE (DESCRIPCION LIKE '%" & ArrCadena(0) & "%')"
                    For index = 0 To ArrCadena.GetUpperBound(0)
                        CADENAsql = CADENAsql & " UNION SELECT DISTINCT DESCRIPCION FROM [dbo].[CAT DE CARGOS DE ESTRUCTURA] WHERE (DESCRIPCION LIKE '%" & ArrCadena(index) & "%')"
                    Next
                End If
                CADENAsql = CADENAsql & "  ORDER BY DESCRIPCION"
                GoTo SALTO
            End If
            If Me.CheckBox1.Checked = True Then
                CADENAsql = "Select DISTINCT DESCRIPCION from [CAT DE CARGOS DE ESTRUCTURA] WHERE (DESCRIPCION LIKE '%" & Me.TextBox3.Text & "%') ORDER BY DESCRIPCION"
            End If
        End If
SALTO:
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[CAT DE CARGOS DE ESTRUCTURA]")
        DGV.DataSource = MiDataSet.Tables("[CAT DE CARGOS DE ESTRUCTURA]")
        Me.Label2.Text = DGV.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW()
        'Me.DGV.Columns(0).HeaderText = "ORDEN"
        'Me.DGV.Columns(0).Width = 0
        'Me.DGV.Columns(0).Visible = False

        'Me.DGV.Columns(1).HeaderText = "D_N1"
        'Me.DGV.Columns(1).Width = 0
        'Me.DGV.Columns(1).Visible = False

        'Me.DGV.Columns(2).HeaderText = "D_N2"
        'Me.DGV.Columns(2).Width = 0
        'Me.DGV.Columns(2).Visible = False

        Me.DGV.Columns(0).HeaderText = "Cargos"
        Me.DGV.Columns(0).Width = 320
        Me.DGV.Columns(0).Visible = True
        Me.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft

        'Me.DGV.Columns(4).HeaderText = "ID_SITUACION"
        'Me.DGV.Columns(4).Width = 0
        'Me.DGV.Columns(4).Visible = False
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
                    maCARGO = Me.DGV.Item(0, I).Value
                    Exit Sub
                Else
                    maCARGO = ""
                End If
            Next
        End If
    End Sub
    Private Sub DGV_DoubleClick(sender As Object, e As EventArgs) Handles DGV.DoubleClick
        If maCARGO <> "" Then
            If ADDUPD = True Then
                If Len(Frm_Aspirantes_Add.TextBox4.Text) = 0 Then
                    Frm_Aspirantes_Add.TextBox4.Text = maCARGO
                Else
                    Frm_Aspirantes_Add.TextBox4.Text = Frm_Aspirantes_Add.TextBox4.Text & ", " & maCARGO
                End If
            End If
            If ADDUPD = False Then
                If Len(Frm_Aspirantes_Upd.TextBox4.Text) = 0 Then
                    Frm_Aspirantes_Upd.TextBox4.Text = maCARGO
                Else
                    Frm_Aspirantes_Upd.TextBox4.Text = Frm_Aspirantes_Upd.TextBox4.Text & ", " & maCARGO
                End If
            End If
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