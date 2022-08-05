Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Cat_Agregar_Causas_Legales
    Private Sub Frm_Cat_Agregar_Causas_Legales_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_MOVIMIENTO_DE_EXP()
        Me.TextBox2.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox2.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub Frm_Cat_Agregar_Causas_Legales_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = ""
            Me.TextBox2.Text = ""
            Me.TextBox2.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_MOVIMIENTO_DE_EXP()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE MOVIMIENTO DE EXP] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_T_M_EXP"
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
    Private Sub GENERAR_CODIGO()
        CODIGO = 0
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            MiDataAdapter = New SqlDataAdapter("SELECT ID_ST_M_L_EXP FROM [CAT DE SUB TIPO DE MOVIMIENTO LEGAL DE EXP] ORDER BY ID_ST_M_L_EXP", Conexion.CxRRHH)
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
            MsgBox("Debe seleccionar el Tipo Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox2.Text) = 0 Then
            MsgBox("Debe digitar la Descripción Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox2.Focus()
            Exit Sub
        End If
        Call GRABAR_REGISTRO()
        c34ID_ST_M_L_EXP = Me.TextBox1.Text
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
        Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        CERRAR = False
        Me.ComboBox1.Focus()
        'Me.Close()
    End Sub
    Private Sub GRABAR_REGISTRO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [CAT DE SUB TIPO DE MOVIMIENTO LEGAL DE EXP] (ID_T_M_EXP, ID_ST_M_L_EXP, DESCRIPCION)" &
                                       "values (" & Me.ComboBox1.SelectedValue & ", " & CInt(Me.TextBox1.Text) & ", '" & Trim(Me.TextBox2.Text) & "')"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter("Select * from [VISTA CAT DE SUB TIPO DE MOVIMIENTO LEGAL DE EXP] ORDER BY D_T_M_EXP, D_ST_M_L_EXP", Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA CAT DE SUB TIPO DE MOVIMIENTO LEGAL DE EXP]")
        Frm_Cat_Causas_Legales.DGV.DataSource = MiDataSet.Tables("[VISTA CAT DE SUB TIPO DE MOVIMIENTO LEGAL DE EXP]")
        Frm_Cat_Causas_Legales.Label2.Text = "Total de Registros: " & Frm_Cat_Causas_Legales.DGV.RowCount
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW()
        Frm_Cat_Causas_Legales.DGV.Columns(0).HeaderText = "ID_T_M_EXP"
        Frm_Cat_Causas_Legales.DGV.Columns(0).Width = 0
        Frm_Cat_Causas_Legales.DGV.Columns(0).Visible = False
        Frm_Cat_Causas_Legales.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Cat_Causas_Legales.DGV.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Frm_Cat_Causas_Legales.DGV.Columns(1).HeaderText = "Tipo Movimiento"
        Frm_Cat_Causas_Legales.DGV.Columns(1).Width = 330
        Frm_Cat_Causas_Legales.DGV.Columns(1).Visible = True
        Frm_Cat_Causas_Legales.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Cat_Causas_Legales.DGV.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Frm_Cat_Causas_Legales.DGV.Columns(2).HeaderText = "Id"
        Frm_Cat_Causas_Legales.DGV.Columns(2).Width = 30
        Frm_Cat_Causas_Legales.DGV.Columns(2).Visible = True
        Frm_Cat_Causas_Legales.DGV.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Cat_Causas_Legales.DGV.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Frm_Cat_Causas_Legales.DGV.Columns(2).DefaultCellStyle.BackColor = Color.Black

        Frm_Cat_Causas_Legales.DGV.Columns(3).HeaderText = "Sub Tipo de Movimiento"
        Frm_Cat_Causas_Legales.DGV.Columns(3).Width = 670
        Frm_Cat_Causas_Legales.DGV.Columns(3).Visible = True
        Frm_Cat_Causas_Legales.DGV.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Cat_Causas_Legales.DGV.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Frm_Cat_Causas_Legales.DGV.Columns(3).DefaultCellStyle.BackColor = Color.LightGreen
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        Dim I As Integer
        For I = 0 To Frm_Cat_Causas_Legales.DGV.RowCount - 1
            If Frm_Cat_Causas_Legales.DGV.Item(2, I).Value = c34ID_ST_M_L_EXP Then
                Frm_Cat_Causas_Legales.DGV.Rows(I).Selected = True
                Frm_Cat_Causas_Legales.DGV.CurrentCell = Frm_Cat_Causas_Legales.DGV.Item(2, I)
                Exit Sub
            End If
        Next
    End Sub
End Class