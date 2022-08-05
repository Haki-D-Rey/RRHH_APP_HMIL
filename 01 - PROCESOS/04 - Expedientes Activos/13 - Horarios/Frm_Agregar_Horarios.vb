Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Agregar_Horarios
    Private Sub Frm_Agregar_Horarios_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = ""
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Agregar_Horarios_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = ""
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_HORARIO()
        Me.ComboBox1.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.TextBox1.Text = ""
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_HORARIO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE HORARIO] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_T_HORARIO"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
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
            MiDataAdapter = New SqlDataAdapter("SELECT ID_H FROM [MAESTRO DE HORARIOS] ORDER BY ID_H", Conexion.CxRRHH)
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
            MsgBox("Debe seleccionar el Tipo de Horario Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        Call GRABAR_REGISTRO()
        vmhiID_H = Me.TextBox1.Text
        Me.TextBox1.Text = ""
        Me.CARGAR_DATAGRIDVIEW_16()
        Me.ARMAR_DATAGRIDVIEW_16()
        Me.SELECCIONAR_REGISTRO_DATAGRIDVIEW_16()
        CERRAR = True
    End Sub
    Dim SQL As String
    Private Sub GRABAR_REGISTRO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()

        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            SQL = "INSERT INTO [MAESTRO DE HORARIOS] (ID_H, ID_M_P, ID_T_HORARIO)" &
                       "values (" & Me.TextBox1.Text & ", " & vID_M_P & ", " & Me.ComboBox1.SelectedValue & ")"
            Dim cmd As New SqlCommand(SQL, CxRRHH)
            cmd.ExecuteNonQuery()
            cmd = Nothing

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_16()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE HORARIOS]")
        Frm_Expedientes_Activos.DGV16.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE HORARIOS]")
        Frm_Expedientes_Activos.Label102.Text = Frm_Expedientes_Activos.DGV16.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_16()
        Frm_Expedientes_Activos.DGV16.Columns(0).HeaderText = "Id"
        Frm_Expedientes_Activos.DGV16.Columns(0).Width = 30
        Frm_Expedientes_Activos.DGV16.Columns(0).Visible = True
        Frm_Expedientes_Activos.DGV16.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV16.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Frm_Expedientes_Activos.DGV16.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Frm_Expedientes_Activos.DGV16.Columns(1).HeaderText = "ID_M_P"
        Frm_Expedientes_Activos.DGV16.Columns(1).Width = 0
        Frm_Expedientes_Activos.DGV16.Columns(1).Visible = False

        Frm_Expedientes_Activos.DGV16.Columns(2).HeaderText = "ID_T_HORARIO"
        Frm_Expedientes_Activos.DGV16.Columns(2).Width = 0
        Frm_Expedientes_Activos.DGV16.Columns(2).Visible = False

        Frm_Expedientes_Activos.DGV16.Columns(3).HeaderText = "Horario"
        Frm_Expedientes_Activos.DGV16.Columns(3).Width = 690
        Frm_Expedientes_Activos.DGV16.Columns(3).Visible = True
        Frm_Expedientes_Activos.DGV16.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV16.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Frm_Expedientes_Activos.DGV16.Columns(4).HeaderText = "Hora Entrada"
        Frm_Expedientes_Activos.DGV16.Columns(4).Width = 120
        Frm_Expedientes_Activos.DGV16.Columns(4).Visible = True
        Frm_Expedientes_Activos.DGV16.Columns(4).DefaultCellStyle.Format = "t"
        Frm_Expedientes_Activos.DGV16.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV16.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV16.Columns(5).HeaderText = "Hora Salida"
        Frm_Expedientes_Activos.DGV16.Columns(5).Width = 120
        Frm_Expedientes_Activos.DGV16.Columns(5).Visible = True
        Frm_Expedientes_Activos.DGV16.Columns(5).DefaultCellStyle.Format = "t"
        Frm_Expedientes_Activos.DGV16.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV16.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV16.Columns(6).HeaderText = "Horas Reales"
        Frm_Expedientes_Activos.DGV16.Columns(6).Width = 150
        Frm_Expedientes_Activos.DGV16.Columns(6).Visible = True
        Frm_Expedientes_Activos.DGV16.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV16.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_16()
        Dim I As Integer
        For I = 0 To Frm_Expedientes_Activos.DGV16.RowCount - 1
            If Frm_Expedientes_Activos.DGV16.Item(0, I).Value = vmhiID_H Then
                Frm_Expedientes_Activos.DGV16.Rows(I).Selected = True
                Frm_Expedientes_Activos.DGV16.CurrentCell = Frm_Expedientes_Activos.DGV16.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
End Class