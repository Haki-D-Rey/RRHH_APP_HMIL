Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Agregar_Idioma
    Private Sub Frm_Agregar_Idioma_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = ""
            Me.CheckBox1.Checked = False
            Me.TextBox2.Text = ""
            Me.CheckBox2.Checked = False
            Me.TextBox3.Text = ""
            Me.CheckBox3.Checked = False
            Me.TextBox4.Text = ""
            Me.TextBox5.Text = ""
            Me.ComboBox1.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Agregar_Idioma_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = ""
        Call CARGAR_COMBO_CAT_DE_IDIOMAS()
        Me.CheckBox1.Checked = False
        Me.TextBox2.Text = ""
        Me.CheckBox2.Checked = False
        Me.TextBox3.Text = ""
        Me.CheckBox3.Checked = False
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Me.ComboBox1.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox1.Text = ""
        Me.CheckBox1.Checked = False
        Me.TextBox2.Text = ""
        Me.CheckBox2.Checked = False
        Me.TextBox3.Text = ""
        Me.CheckBox3.Checked = False
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Me.ComboBox1.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_IDIOMAS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE IDIOMAS] order by ID_IDIOMA", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_IDIOMA"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub CheckBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub CheckBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub CheckBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub GENERAR_CODIGO()
        CODIGO = 0
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            MiDataAdapter = New SqlDataAdapter("SELECT ID_M_I FROM [MAESTRO DE IDIOMAS] ORDER BY ID_M_I", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                For I = 0 To MiDataTable.Rows.Count - 1
                    CODIGO = MiDataTable.Rows(I).Item(0).ToString
                Next
                CODIGO += 1
            Else
                CODIGO = 1
            End If
            Me.TextBox1.Text = CODIGO
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.TextBox1.Text = "0"
        Call GENERAR_CODIGO()
        If TextBox1.Text = "0" Or TextBox1.Text = "" Then
            MsgBox("No se ha generado el Codigo del registro", vbInformation, "Mensaje del Sistema")
            Me.TextBox1.Text = "0"
            Me.Button2.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Idioma Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If (Me.CheckBox1.Checked = True And Len(Me.TextBox2.Text) = 0) Or (Me.CheckBox1.Checked = True And Me.TextBox2.Text = "0") Then
            MsgBox("Si no hay Porcentaje existente debe desmarcar la casilla <HABLA>", vbInformation, "Mensaje del Sistema")
            Me.TextBox2.Focus()
            Exit Sub
        End If
        If (Me.CheckBox2.Checked = True And Len(Me.TextBox3.Text) = 0) Or (Me.CheckBox2.Checked = True And Me.TextBox3.Text = "0") Then
            MsgBox("Si no hay Porcentaje existente debe desmarcar la casilla <LEE>", vbInformation, "Mensaje del Sistema")
            Me.TextBox3.Focus()
            Exit Sub
        End If
        If (Me.CheckBox3.Checked = True And Len(Me.TextBox4.Text) = 0) Or (Me.CheckBox3.Checked = True And Me.TextBox4.Text = "0") Then
            MsgBox("Si no hay Porcentaje existente debe desmarcar la casilla <ESCRIBE>", vbInformation, "Mensaje del Sistema")
            Me.TextBox4.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox5.Text) = 0 Then
            Me.TextBox5.Text = "."
        End If
        Call GRABAR_REGISTRO()
        vmiID_M_I = Me.TextBox1.Text
        CADENAsql = "SELECT * FROM [VISTA MAESTRO DE IDIOMAS] WHERE ID_M_P = " & Frm_Expedientes_Activos.TextBox35.Text & " ORDER BY ID_M_P, ID_M_I"
        Call CARGAR_DATAGRIDVIEW_6()
        Call ARMAR_DATAGRIDVIEW_6()
        Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_6()
        CERRAR = False
    End Sub
    Private Sub GRABAR_REGISTRO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE IDIOMAS] (ID_M_I, ID_M_P, ID_IDIOMA, HABLA, P_HABLA, LEE, PORCENTAJE_L, ESCRIBE, PORCENTAJE_E, OBSERVACIONES)" &
                                  "values (" & CInt(Me.TextBox1.Text) & ", " & Frm_Expedientes_Activos.TextBox35.Text & ", " & Me.ComboBox1.SelectedValue & ", '" & Me.CheckBox1.Checked & "', '" & Me.TextBox2.Text & "', '" & Me.CheckBox2.Checked & "', '" & Me.TextBox3.Text & "', '" & Me.CheckBox3.Checked & "', '" & Me.TextBox4.Text & "', '" & Me.TextBox5.Text & "')"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_6()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE IDIOMAS]")
        Frm_Expedientes_Activos.DGV6.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE IDIOMAS]")
        Frm_Expedientes_Activos.Label84.Text = Frm_Expedientes_Activos.DGV6.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_6()
        Frm_Expedientes_Activos.DGV6.Columns(0).HeaderText = "Id"
        Frm_Expedientes_Activos.DGV6.Columns(0).Width = 30
        Frm_Expedientes_Activos.DGV6.Columns(0).Visible = True
        Frm_Expedientes_Activos.DGV6.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV6.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Frm_Expedientes_Activos.DGV6.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Frm_Expedientes_Activos.DGV6.Columns(1).HeaderText = "ID_M_P"
        Frm_Expedientes_Activos.DGV6.Columns(1).Width = 0
        Frm_Expedientes_Activos.DGV6.Columns(1).Visible = False

        Frm_Expedientes_Activos.DGV6.Columns(2).HeaderText = "ID_IDIOMA"
        Frm_Expedientes_Activos.DGV6.Columns(2).Width = 0
        Frm_Expedientes_Activos.DGV6.Columns(2).Visible = False

        Frm_Expedientes_Activos.DGV6.Columns(3).HeaderText = "Idioma"
        Frm_Expedientes_Activos.DGV6.Columns(3).Width = 160
        Frm_Expedientes_Activos.DGV6.Columns(3).Visible = True
        Frm_Expedientes_Activos.DGV6.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV6.Columns(4).HeaderText = "Habla"
        Frm_Expedientes_Activos.DGV6.Columns(4).Width = 70
        Frm_Expedientes_Activos.DGV6.Columns(4).Visible = True
        Frm_Expedientes_Activos.DGV6.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV6.Columns(5).HeaderText = "%"
        Frm_Expedientes_Activos.DGV6.Columns(5).Width = 70
        Frm_Expedientes_Activos.DGV6.Columns(5).Visible = True
        Frm_Expedientes_Activos.DGV6.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV6.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV6.Columns(6).HeaderText = "Lee"
        Frm_Expedientes_Activos.DGV6.Columns(6).Width = 70
        Frm_Expedientes_Activos.DGV6.Columns(6).Visible = True
        Frm_Expedientes_Activos.DGV6.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV6.Columns(7).HeaderText = "%"
        Frm_Expedientes_Activos.DGV6.Columns(7).Width = 70
        Frm_Expedientes_Activos.DGV6.Columns(7).Visible = True
        Frm_Expedientes_Activos.DGV6.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV6.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV6.Columns(8).HeaderText = "Escribe"
        Frm_Expedientes_Activos.DGV6.Columns(8).Width = 70
        Frm_Expedientes_Activos.DGV6.Columns(8).Visible = True
        Frm_Expedientes_Activos.DGV6.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV6.Columns(9).HeaderText = "%"
        Frm_Expedientes_Activos.DGV6.Columns(9).Width = 70
        Frm_Expedientes_Activos.DGV6.Columns(9).Visible = True
        Frm_Expedientes_Activos.DGV6.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV6.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV6.Columns(10).HeaderText = "Observaciones"
        Frm_Expedientes_Activos.DGV6.Columns(10).Width = 500
        Frm_Expedientes_Activos.DGV6.Columns(10).Visible = True
        Frm_Expedientes_Activos.DGV6.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_6()
        Dim I As Integer
        For I = 0 To Frm_Expedientes_Activos.DGV6.RowCount - 1
            If Frm_Expedientes_Activos.DGV6.Item(0, I).Value = vmiID_M_I Then
                Frm_Expedientes_Activos.DGV6.Rows(I).Selected = True
                Frm_Expedientes_Activos.DGV6.CurrentCell = Frm_Expedientes_Activos.DGV6.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
End Class