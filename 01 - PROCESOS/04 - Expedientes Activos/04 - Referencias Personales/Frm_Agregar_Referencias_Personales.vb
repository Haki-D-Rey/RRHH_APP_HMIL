Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Public Class Frm_Agregar_Referencias_Personales
    Private Sub Frm_Agregar_Referencias_Personales_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = "0"
        Me.TextBox2.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = "0"
        Me.TextBox2.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub Frm_Agregar_Referencias_Personales_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = ""
            Me.TextBox2.Text = ""
            Me.TextBox3.Text = ""
            Me.TextBox4.Text = ""
            Me.TextBox5.Text = "0"
            Me.TextBox2.Focus()
            CERRAR = True
            Me.Close()
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
    Private Sub GENERAR_CODIGO()
        CODIGO = 0
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            MiDataAdapter = New SqlDataAdapter("SELECT ID_RP FROM [MAESTRO DE REFERENCIAS PERSONALES] ORDER BY ID_RP", Conexion.CxRRHH)
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
    Dim vcMIXTA, vcMILITAR, vcPAME As Boolean
    Private Sub BUSCAR_ESTRUCTURA_CARGO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [MAESTRO DE CARGOS] WHERE ID_M_C = '" & Frm_Expedientes_Activos.TextBox36.Text & "'", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                vcMIXTA = MiDataTable.Rows(0).Item(15).ToString
                vcMILITAR = MiDataTable.Rows(0).Item(16).ToString
                vcPAME = MiDataTable.Rows(0).Item(17).ToString
            Else
                vcMIXTA = False
                vcMILITAR = False
                vcPAME = False
            End If
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
        If Len(Me.TextBox2.Text) = 0 Then
            MsgBox("Debe digitar los Apellidos y Nombres correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox3.Text) = 0 Then
            MsgBox("Debe digitar la Dirección correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox3.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox4.Text) = 0 Then
            MsgBox("Debe digitar la Ocupación correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox4.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox5.Text) = 0 Then
            MsgBox("Debe digitar el Número Telefonico correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox5.Focus()
            Exit Sub
        End If
        Call GRABAR_REGISTRO()
        vmrID_RP = Me.TextBox1.Text
        CADENAsql = "SELECT * FROM [MAESTRO DE REFERENCIAS PERSONALES] WHERE ID_M_P = " & Frm_Expedientes_Activos.TextBox35.Text & " ORDER BY ID_M_P, ID_RP"
        Call CARGAR_DATAGRIDVIEW_3()
        Call ARMAR_DATAGRIDVIEW_3()
        Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_3()
        CERRAR = False
    End Sub
    Private Sub GRABAR_REGISTRO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE REFERENCIAS PERSONALES] (ID_RP, ID_M_P, APELLIDOS_NOMBRES, DIRECCION, OCUPACION, TELEFONO)" &
                                  "values (" & CInt(Me.TextBox1.Text) & ", " & Frm_Expedientes_Activos.TextBox35.Text & ", '" & Me.TextBox2.Text & "', '" & Me.TextBox3.Text & "', '" & Me.TextBox4.Text & "', '" & Me.TextBox5.Text & "')"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_3()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[MAESTRO DE REFERENCIAS PERSONALES]")
        Frm_Expedientes_Activos.DGV3.DataSource = MiDataSet.Tables("[MAESTRO DE REFERENCIAS PERSONALES]")
        Frm_Expedientes_Activos.Label81.Text = Frm_Expedientes_Activos.DGV3.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_3()
        Frm_Expedientes_Activos.DGV3.Columns(0).HeaderText = "Id"
        Frm_Expedientes_Activos.DGV3.Columns(0).Width = 30
        Frm_Expedientes_Activos.DGV3.Columns(0).Visible = True
        Frm_Expedientes_Activos.DGV3.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Activos.DGV3.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Frm_Expedientes_Activos.DGV3.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Frm_Expedientes_Activos.DGV3.Columns(1).HeaderText = "ID_M_P"
        Frm_Expedientes_Activos.DGV3.Columns(1).Width = 0
        Frm_Expedientes_Activos.DGV3.Columns(1).Visible = False

        Frm_Expedientes_Activos.DGV3.Columns(2).HeaderText = "Apellidos y Nombres"
        Frm_Expedientes_Activos.DGV3.Columns(2).Width = 310
        Frm_Expedientes_Activos.DGV3.Columns(2).Visible = True
        Frm_Expedientes_Activos.DGV3.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV3.Columns(3).HeaderText = "Direccion"
        Frm_Expedientes_Activos.DGV3.Columns(3).Width = 420
        Frm_Expedientes_Activos.DGV3.Columns(3).Visible = True
        Frm_Expedientes_Activos.DGV3.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV3.Columns(4).HeaderText = "Ocupacion"
        Frm_Expedientes_Activos.DGV3.Columns(4).Width = 250
        Frm_Expedientes_Activos.DGV3.Columns(4).Visible = True
        Frm_Expedientes_Activos.DGV3.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Activos.DGV3.Columns(5).HeaderText = "Telefono"
        Frm_Expedientes_Activos.DGV3.Columns(5).Width = 100
        Frm_Expedientes_Activos.DGV3.Columns(5).Visible = True
        Frm_Expedientes_Activos.DGV3.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_3()
        Dim I As Integer
        For I = 0 To Frm_Expedientes_Activos.DGV3.RowCount - 1
            If Frm_Expedientes_Activos.DGV3.Item(0, I).Value = vmrID_RP Then
                Frm_Expedientes_Activos.DGV3.Rows(I).Selected = True
                Frm_Expedientes_Activos.DGV3.CurrentCell = Frm_Expedientes_Activos.DGV3.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
End Class