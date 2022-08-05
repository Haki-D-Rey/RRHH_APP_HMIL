Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Cat_Agregar_Municipio_D
    Private Sub Frm_Cat_Agregar_Municipio_D_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Call CARGAR_COMBO_CAT_DE_DPTO_D()
        Me.TextBox2.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox2.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub Frm_Cat_Agregar_Municipio_D_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = ""
            Me.TextBox2.Text = ""
            Me.TextBox2.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_DPTO_D()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE DPTO_D] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_DPTO_D"
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
            MiDataAdapter = New SqlDataAdapter("SELECT ID_M_D FROM [CAT DE MUNICIPIO_D] ORDER BY ID_M_D", Conexion.CxRRHH)
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
            MsgBox("Debe seleccionar el Departamento Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox2.Text) = 0 Then
            MsgBox("Debe digitar la Descripción Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox2.Focus()
            Exit Sub
        End If
        Call GRABAR_REGISTRO()
        c16ID_M_D = Me.TextBox1.Text
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
        Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        CERRAR = False
        Me.ComboBox1.Focus()
        'Me.Close()
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter("Select * from [VISTA CAT DE MUNICIPIO_D] ORDER BY D_DPTO_D, DESCRIPCION", Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA CAT DE MUNICIPIO_D]")
        Frm_Cat_Municipio_D.DGV.DataSource = MiDataSet.Tables("[VISTA CAT DE MUNICIPIO_D]")
        Frm_Cat_Municipio_D.Label2.Text = "Total de Registros: " & Frm_Cat_Municipio_D.DGV.RowCount
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW()
        Frm_Cat_Municipio_D.DGV.Columns(0).HeaderText = "ID_DPTO_D"
        Frm_Cat_Municipio_D.DGV.Columns(0).Width = 0
        Frm_Cat_Municipio_D.DGV.Columns(0).Visible = False
        Frm_Cat_Municipio_D.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Cat_Municipio_D.DGV.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Frm_Cat_Municipio_D.DGV.Columns(1).HeaderText = "Departamento"
        Frm_Cat_Municipio_D.DGV.Columns(1).Width = 330
        Frm_Cat_Municipio_D.DGV.Columns(1).Visible = True
        Frm_Cat_Municipio_D.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Cat_Municipio_D.DGV.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Frm_Cat_Municipio_D.DGV.Columns(2).HeaderText = "Id"
        Frm_Cat_Municipio_D.DGV.Columns(2).Width = 30
        Frm_Cat_Municipio_D.DGV.Columns(2).Visible = True
        Frm_Cat_Municipio_D.DGV.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Cat_Municipio_D.DGV.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Frm_Cat_Municipio_D.DGV.Columns(2).DefaultCellStyle.BackColor = Color.Black

        Frm_Cat_Municipio_D.DGV.Columns(3).HeaderText = "Descripcion"
        Frm_Cat_Municipio_D.DGV.Columns(3).Width = 670
        Frm_Cat_Municipio_D.DGV.Columns(3).Visible = True
        Frm_Cat_Municipio_D.DGV.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Cat_Municipio_D.DGV.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Frm_Cat_Municipio_D.DGV.Columns(3).DefaultCellStyle.BackColor = Color.LightGreen
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        Dim I As Integer
        For I = 0 To Frm_Cat_Municipio_D.DGV.RowCount - 1
            If Frm_Cat_Municipio_D.DGV.Item(2, I).Value = c16ID_M_D Then
                Frm_Cat_Municipio_D.DGV.Rows(I).Selected = True
                Frm_Cat_Municipio_D.DGV.CurrentCell = Frm_Cat_Municipio_D.DGV.Item(2, I)
                Exit Sub
            End If
        Next
    End Sub
    Private Sub GRABAR_REGISTRO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [CAT DE MUNICIPIO_D] (ID_DPTO_D, ID_M_D, DESCRIPCION)" &
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
End Class