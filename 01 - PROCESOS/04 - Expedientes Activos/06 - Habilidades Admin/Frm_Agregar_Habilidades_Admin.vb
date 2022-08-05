Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Agregar_Habilidades_Admin
    Private Sub Frm_Agregar_Habilidades_Admin_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = ""
            Me.TextBox2.Text = "0"
            Me.TextBox3.Text = ""
            Me.ComboBox1.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Agregar_Habilidades_Admin_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = ""
        Call CARGAR_COMBO_CAT_DE_HABILIDADES_ADMINISTRATIVAS()
        Me.TextBox2.Text = "0"
        Me.TextBox3.Text = ""
        Me.ComboBox1.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = "0"
        Me.TextBox3.Text = ""
        Me.ComboBox1.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_HABILIDADES_ADMINISTRATIVAS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE HABILIDADES ADMINISTRATIVAS] order by ID_HABILIDAD", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_HABILIDAD"
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
    Private Sub GENERAR_CODIGO()
        CODIGO = 0
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            MiDataAdapter = New SqlDataAdapter("SELECT ID_H_ADMIN FROM [MAESTRO DE HABILIDADES ADMIN] ORDER BY ID_H_ADMIN", Conexion.CxRRHH)
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
            MsgBox("Debe seleccionar la Habilidad Administrativa Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If Val(Me.TextBox2.Text) < 10 Then
            MsgBox("Debe digitar el porcentaje (%) de control de la Habilidad Administrativa seleccionda", vbInformation, "Mensaje del Sistema")
            Me.TextBox2.Focus()
            Exit Sub
        End If
        Call GRABAR_REGISTRO()
        vmhID_H_ADMIN = Me.TextBox1.Text
        CADENAsql = "SELECT * FROM [VISTA MAESTRO DE HABILIDADES ADMIN] WHERE ID_M_P = " & Frm_Expedientes_Activos.TextBox35.Text & " ORDER BY ID_M_P, ID_H_ADMIN"
        Call CARGAR_DATAGRIDVIEW_5()
        Call ARMAR_DATAGRIDVIEW_5()
        Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_5()
        CERRAR = False
    End Sub
    Private Sub GRABAR_REGISTRO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE HABILIDADES ADMIN] (ID_H_ADMIN, ID_M_P, ID_HABILIDAD, PORCENTAJE, OBSERVACIONES)" &
                                  "values (" & CInt(Me.TextBox1.Text) & ", " & Frm_Expedientes_Activos.TextBox35.Text & ", " & Me.ComboBox1.SelectedValue & ", '" & Me.TextBox2.Text & "', '" & Me.TextBox3.Text & "')"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_5()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE HABILIDADES ADMIN]")
        Frm_Expedientes_Activos.DGV5.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE HABILIDADES ADMIN]")
        Frm_Expedientes_Activos.Label83.Text = Frm_Expedientes_Activos.DGV5.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_5()
        Frm_Expedientes_Activos.DGV5.Columns(0).HeaderText = "Id"
        Frm_Expedientes_Activos.DGV5.Columns(0).Width = 30
        Frm_Expedientes_Activos.DGV5.Columns(0).Visible = True
        Frm_Expedientes_Activos.DGV5.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV5.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Frm_Expedientes_Activos.DGV5.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Frm_Expedientes_Activos.DGV5.Columns(1).HeaderText = "ID_M_P"
        Frm_Expedientes_Activos.DGV5.Columns(1).Width = 0
        Frm_Expedientes_Activos.DGV5.Columns(1).Visible = False

        Frm_Expedientes_Activos.DGV5.Columns(2).HeaderText = "ID_HABILIDAD"
        Frm_Expedientes_Activos.DGV5.Columns(2).Width = 0
        Frm_Expedientes_Activos.DGV5.Columns(2).Visible = False

        Frm_Expedientes_Activos.DGV5.Columns(3).HeaderText = "Habilidad"
        Frm_Expedientes_Activos.DGV5.Columns(3).Width = 200
        Frm_Expedientes_Activos.DGV5.Columns(3).Visible = True
        Frm_Expedientes_Activos.DGV5.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV5.Columns(4).HeaderText = "Porcentaje (%)"
        Frm_Expedientes_Activos.DGV5.Columns(4).Width = 80
        Frm_Expedientes_Activos.DGV5.Columns(4).Visible = True
        Frm_Expedientes_Activos.DGV5.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV5.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV5.Columns(5).HeaderText = "Observaciones"
        Frm_Expedientes_Activos.DGV5.Columns(5).Width = 800
        Frm_Expedientes_Activos.DGV5.Columns(5).Visible = True
        Frm_Expedientes_Activos.DGV5.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_5()
        Dim I As Integer
        For I = 0 To Frm_Expedientes_Activos.DGV5.RowCount - 1
            If Frm_Expedientes_Activos.DGV5.Item(0, I).Value = vmhID_H_ADMIN Then
                Frm_Expedientes_Activos.DGV5.Rows(I).Selected = True
                Frm_Expedientes_Activos.DGV5.CurrentCell = Frm_Expedientes_Activos.DGV5.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
End Class