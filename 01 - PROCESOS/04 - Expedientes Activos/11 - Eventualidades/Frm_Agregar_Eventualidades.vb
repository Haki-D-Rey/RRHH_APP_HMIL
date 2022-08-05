Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Agregar_Eventualidades
    Private Sub Frm_Agregar_Eventualidades_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = ""
            Me.TextBox2.Text = ""
            Me.DateTimePicker1.Value = Now
            Me.TextBox3.Text = ""
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Agregar_Eventualidades_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.DateTimePicker1.Value = Now
        Call CARGAR_COMBO_CAT_DE_EVENTUALIDADES()
        Call CARGAR_COMBO_CAT_DE_CAUSAS_REALES_DE_EVENTOS
        Me.TextBox3.Text = ""
        Me.TextBox2.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.DateTimePicker1.Value = Now
        Me.TextBox3.Text = ""
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_EVENTUALIDADES()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE EVENTUALIDADES] order by ID_E", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox3
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_E"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_CAUSAS_REALES_DE_EVENTOS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE CAUSAS REALES DE EVENTOS]  order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count <> 0 Then
                With Me.ComboBox4
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_CR_EVENTOS"
                End With
            Else
                Me.ComboBox4.Text = ""
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ComboBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox3.KeyDown
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
    Private Sub DateTimePicker1_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker1.KeyDown
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
            MiDataAdapter = New SqlDataAdapter("SELECT ID_M_EV FROM [MAESTRO DE EVENTUALIDADES] ORDER BY ID_M_EV", Conexion.CxRRHH)
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
        If Len(Me.TextBox2.Text) < 5 Then
            MsgBox("Debe digitar Correctamente los Nombres y Apellidos del Jefe Inmediato", vbInformation, "Mensaje del Sistema")
            Me.TextBox2.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox3.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar la Eventualidad Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox3.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox4.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar la Causa Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox4.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox3.Text) = 0 Then
            Me.TextBox3.Text = ""
        End If

        Call GRABAR_REGISTRO()
        vmeID_M_EV = Me.TextBox1.Text
        CADENAsql = "SELECT * FROM [VISTA MAESTRO DE EVENTUALIDADES] WHERE ID_M_P = " & Frm_Expedientes_Activos.TextBox35.Text & " ORDER BY FEC_EVE DESC"
        Call CARGAR_DATAGRIDVIEW_10()
        Call ARMAR_DATAGRIDVIEW_10()
        Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_10()
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.DateTimePicker1.Value = Now
        Me.TextBox3.Text = ""
        Me.TextBox2.Focus()
    End Sub
    Private Sub GRABAR_REGISTRO()
        Dim fFEC_EVE As String = Me.DateTimePicker1.Text
        Dim fechaEventualidad = "Convert(VARCHAR(10), Convert(Date, '" + fFEC_EVE + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE EVENTUALIDADES] (ID_M_EV, ID_M_P, JEFE, FEC_EVE, ID_CR_EVENTOS, ID_E, EVENTUALIDAD)" &
                                  "values (" & CInt(Me.TextBox1.Text) & ", " & Frm_Expedientes_Activos.TextBox35.Text & ", '" & Me.TextBox2.Text & "', " & fechaEventualidad & ", " & Me.ComboBox4.SelectedValue & ", " & Me.ComboBox3.SelectedValue & ", '" & Me.TextBox3.Text & "')"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_10()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE EVENTUALIDADES]")
        Frm_Expedientes_Activos.DGV10.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE EVENTUALIDADES]")
        Frm_Expedientes_Activos.Label88.Text = Frm_Expedientes_Activos.DGV10.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_10()
        Frm_Expedientes_Activos.DGV10.Columns(0).HeaderText = "Id"
        Frm_Expedientes_Activos.DGV10.Columns(0).Width = 30
        Frm_Expedientes_Activos.DGV10.Columns(0).Visible = True
        Frm_Expedientes_Activos.DGV10.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV10.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Frm_Expedientes_Activos.DGV10.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Frm_Expedientes_Activos.DGV10.Columns(1).HeaderText = "ID_M_P"
        Frm_Expedientes_Activos.DGV10.Columns(1).Width = 0
        Frm_Expedientes_Activos.DGV10.Columns(1).Visible = False

        Frm_Expedientes_Activos.DGV10.Columns(2).HeaderText = "Nivel 1"
        Frm_Expedientes_Activos.DGV10.Columns(2).Width = 0
        Frm_Expedientes_Activos.DGV10.Columns(2).Visible = False
        Frm_Expedientes_Activos.DGV10.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV10.Columns(3).HeaderText = "Nivel 2"
        Frm_Expedientes_Activos.DGV10.Columns(3).Width = 0
        Frm_Expedientes_Activos.DGV10.Columns(3).Visible = 0
        Frm_Expedientes_Activos.DGV10.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV10.Columns(4).HeaderText = "Jefe Inmediato"
        Frm_Expedientes_Activos.DGV10.Columns(4).Width = 100
        Frm_Expedientes_Activos.DGV10.Columns(4).Visible = True
        Frm_Expedientes_Activos.DGV10.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV10.Columns(5).HeaderText = "Fecha"
        Frm_Expedientes_Activos.DGV10.Columns(5).Width = 70
        Frm_Expedientes_Activos.DGV10.Columns(5).Visible = True
        Frm_Expedientes_Activos.DGV10.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV10.Columns(6).HeaderText = "ID_CR_EVENTOS"
        Frm_Expedientes_Activos.DGV10.Columns(6).Width = 0
        Frm_Expedientes_Activos.DGV10.Columns(6).Visible = False

        Frm_Expedientes_Activos.DGV10.Columns(7).HeaderText = "Causa"
        Frm_Expedientes_Activos.DGV10.Columns(7).Width = 100
        Frm_Expedientes_Activos.DGV10.Columns(7).Visible = True
        Frm_Expedientes_Activos.DGV10.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV10.Columns(8).HeaderText = "ID_E"
        Frm_Expedientes_Activos.DGV10.Columns(8).Width = 0
        Frm_Expedientes_Activos.DGV10.Columns(8).Visible = False

        Frm_Expedientes_Activos.DGV10.Columns(9).HeaderText = "Eventualidad"
        Frm_Expedientes_Activos.DGV10.Columns(9).Width = 140
        Frm_Expedientes_Activos.DGV10.Columns(9).Visible = True
        Frm_Expedientes_Activos.DGV10.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV10.Columns(10).HeaderText = "Observaciones"
        Frm_Expedientes_Activos.DGV10.Columns(10).Width = 670
        Frm_Expedientes_Activos.DGV10.Columns(10).Visible = True
        Frm_Expedientes_Activos.DGV10.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_10()
        Dim I As Integer
        For I = 0 To Frm_Expedientes_Activos.DGV10.RowCount - 1
            If Frm_Expedientes_Activos.DGV10.Item(0, I).Value = vmeID_M_EV Then
                Frm_Expedientes_Activos.DGV10.Rows(I).Selected = True
                Frm_Expedientes_Activos.DGV10.CurrentCell = Frm_Expedientes_Activos.DGV10.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub ComboBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
End Class