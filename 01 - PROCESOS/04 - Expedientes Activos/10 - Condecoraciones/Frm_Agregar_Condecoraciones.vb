Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Agregar_Condecoraciones
    Private Sub Frm_Agregar_Condecoraciones_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = ""
            Call Clases.BUSCAR_FECHA_Y_HORA()
            Me.DateTimePicker1.Value = FECHA_SERVIDOR
            Me.TextBox3.Text = ""
            Me.TextBox4.Text = ""
            Me.ComboBox1.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Agregar_Condecoraciones_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = ""
        Call CARGAR_COMBO_CAT_DE_CONDECORACIONES()
        Call Clases.BUSCAR_FECHA_Y_HORA()
        Me.DateTimePicker1.Value = FECHA_SERVIDOR
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.ComboBox1.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox1.Text = ""
        Call Clases.BUSCAR_FECHA_Y_HORA()
        Me.DateTimePicker1.Value = FECHA_SERVIDOR
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.ComboBox1.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_CONDECORACIONES()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE CONDECORACIONES] order by ID_CONDE", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_CONDE"
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
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs)
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
    Private Sub GENERAR_CODIGO()
        CODIGO = 0
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            MiDataAdapter = New SqlDataAdapter("SELECT ID_M_CONDE FROM [MAESTRO DE CONDECORACIONES] ORDER BY ID_M_CONDE", Conexion.CxRRHH)
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
            MsgBox("Debe seleccionar la Condecoración Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox3.Text) = 0 Then
            Me.TextBox3.Text = ""
        End If
        If Len(Me.TextBox4.Text) = 0 Then
            Me.TextBox4.Text = ""
        End If
        Call GRABAR_REGISTRO()
        vmcID_M_CONDE = Me.TextBox1.Text
        CADENAsql = "SELECT * FROM [VISTA MAESTRO DE CONDECORACIONES] WHERE ID_M_P = " & Frm_Expedientes_Activos.TextBox35.Text & " ORDER BY ID_M_P, ID_M_CONDE"
        Call CARGAR_DATAGRIDVIEW_9()
        Call ARMAR_DATAGRIDVIEW_9()
        Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_9()
        Me.TextBox1.Text = "0"
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.ComboBox1.Focus()
        CERRAR = False
    End Sub
    Dim fechaCONDE As String
    Dim fFEC_CONDE As String
    Private Sub GRABAR_REGISTRO()
        fFEC_CONDE = Mid(Me.DateTimePicker1.Value, 1, 10)
        fechaCONDE = "Convert(VARCHAR(10), Convert(Date, '" + fFEC_CONDE + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE CONDECORACIONES] (ID_M_CONDE, ID_M_P, ID_CONDE, CONDE_A, MOTIVOS, OBSERVACIONES)" &
                                  "values (" & CInt(Me.TextBox1.Text) & ", " & Frm_Expedientes_Activos.TextBox35.Text & ", " & Me.ComboBox1.SelectedValue & ", " & fechaCONDE & ", '" & Me.TextBox3.Text & "', '" & Me.TextBox4.Text & "')"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_9()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE CONDECORACIONES]")
        Frm_Expedientes_Activos.DGV9.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE CONDECORACIONES]")
        Frm_Expedientes_Activos.Label87.Text = Frm_Expedientes_Activos.DGV9.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_9()
        Frm_Expedientes_Activos.DGV9.Columns(0).HeaderText = "ID_C_CONDE"
        Frm_Expedientes_Activos.DGV9.Columns(0).Width = 0
        Frm_Expedientes_Activos.DGV9.Columns(0).Visible = False

        Frm_Expedientes_Activos.DGV9.Columns(1).HeaderText = "Clasificacion"
        Frm_Expedientes_Activos.DGV9.Columns(1).Width = 250
        Frm_Expedientes_Activos.DGV9.Columns(1).Visible = True
        Frm_Expedientes_Activos.DGV9.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV9.Columns(2).HeaderText = "ID_CONDE"
        Frm_Expedientes_Activos.DGV9.Columns(2).Width = 0
        Frm_Expedientes_Activos.DGV9.Columns(2).Visible = False

        Frm_Expedientes_Activos.DGV9.Columns(3).HeaderText = "Codigo"
        Frm_Expedientes_Activos.DGV9.Columns(3).Width = 80
        Frm_Expedientes_Activos.DGV9.Columns(3).Visible = True
        Frm_Expedientes_Activos.DGV9.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV9.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        'Frm_Expedientes_Activos.DGV9.Columns(3).Frozen = True

        Frm_Expedientes_Activos.DGV9.Columns(4).HeaderText = "Condecoracion"
        Frm_Expedientes_Activos.DGV9.Columns(4).Width = 320
        Frm_Expedientes_Activos.DGV9.Columns(4).Visible = True
        Frm_Expedientes_Activos.DGV9.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV9.Columns(5).HeaderText = "Id"
        Frm_Expedientes_Activos.DGV9.Columns(5).Width = 30
        Frm_Expedientes_Activos.DGV9.Columns(5).Visible = True
        Frm_Expedientes_Activos.DGV9.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV9.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Frm_Expedientes_Activos.DGV9.Columns(5).DefaultCellStyle.BackColor = Color.Black

        Frm_Expedientes_Activos.DGV9.Columns(6).HeaderText = "ID_M_P"
        Frm_Expedientes_Activos.DGV9.Columns(6).Width = 0
        Frm_Expedientes_Activos.DGV9.Columns(6).Visible = False

        Frm_Expedientes_Activos.DGV9.Columns(7).HeaderText = "Año de Entrega"
        Frm_Expedientes_Activos.DGV9.Columns(7).Width = 100
        Frm_Expedientes_Activos.DGV9.Columns(7).Visible = True
        Frm_Expedientes_Activos.DGV9.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV9.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV9.Columns(8).HeaderText = "Motivos"
        Frm_Expedientes_Activos.DGV9.Columns(8).Width = 120
        Frm_Expedientes_Activos.DGV9.Columns(8).Visible = True
        Frm_Expedientes_Activos.DGV9.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV9.Columns(9).HeaderText = "Observaciones"
        Frm_Expedientes_Activos.DGV9.Columns(9).Width = 200
        Frm_Expedientes_Activos.DGV9.Columns(9).Visible = True
        Frm_Expedientes_Activos.DGV9.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_9()
        Dim I As Integer
        For I = 0 To Frm_Expedientes_Activos.DGV9.RowCount - 1
            If Frm_Expedientes_Activos.DGV9.Item(5, I).Value = vmcID_M_CONDE Then
                Frm_Expedientes_Activos.DGV9.Rows(I).Selected = True
                Frm_Expedientes_Activos.DGV9.CurrentCell = Frm_Expedientes_Activos.DGV9.Item(5, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub DateTimePicker1_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
End Class