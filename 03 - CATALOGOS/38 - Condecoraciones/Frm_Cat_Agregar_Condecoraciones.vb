Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Cat_Agregar_Condecoraciones
    Private Sub Frm_Cat_Agregar_Condecoraciones_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = ""
            Me.TextBox3.Text = ""
            Me.TextBox4.Text = ""
            Me.TextBox3.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Cat_Agregar_Condecoraciones_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = ""
        Call CARGAR_COMBO_CAT_DE_CLASIFICACION_DE_CONECORACIONES()
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox3.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox1.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox3.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_CLASIFICACION_DE_CONECORACIONES()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE CLASIFICACION DE CONECORACIONES] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox2
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_C_CONDE"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
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
            MiDataAdapter = New SqlDataAdapter("SELECT ID_CONDE FROM [CAT DE CONDECORACIONES] ORDER BY ID_CONDE", Conexion.CxRRHH)
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
        If Len(Me.ComboBox2.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar la Clasificacion", vbInformation, "Mensaje del Sistema")
            Me.ComboBox2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox3.Text) = 0 Then
            MsgBox("Debe digitar la Sigla Correctamente o escribir (.) si no desea una Sigla para este registro", vbInformation, "Mensaje del Sistema")
            Me.TextBox4.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox4.Text) = 0 Then
            MsgBox("Debe digitar la Descripción Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox3.Focus()
            Exit Sub
        End If
        Call GRABAR_REGISTRO()
        c32ID_CONDE = Me.TextBox1.Text
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
        Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        Me.TextBox1.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox3.Focus()
        Me.TextBox3.SelectAll()
        CERRAR = False
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter("Select * from [VISTA CAT DE CONDECORACIONES] ORDER BY ID_C_CONDE, ID_CONDE", Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA CAT DE CONDECORACIONES]")
        Frm_Cat_Condecoraciones.DGV.DataSource = MiDataSet.Tables("[VISTA CAT DE CONDECORACIONES]")
        Frm_Cat_Condecoraciones.Label2.Text = "Total de Registros: " & Frm_Cat_Condecoraciones.DGV.RowCount
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW()
        Frm_Cat_Condecoraciones.DGV.Columns(0).HeaderText = "ID_C_CONDE"
        Frm_Cat_Condecoraciones.DGV.Columns(0).Width = 0
        Frm_Cat_Condecoraciones.DGV.Columns(0).Visible = False

        Frm_Cat_Condecoraciones.DGV.Columns(1).HeaderText = "Clasificacion"
        Frm_Cat_Condecoraciones.DGV.Columns(1).Width = 200
        Frm_Cat_Condecoraciones.DGV.Columns(1).Visible = True
        Frm_Cat_Condecoraciones.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Cat_Condecoraciones.DGV.Columns(2).HeaderText = "Id"
        Frm_Cat_Condecoraciones.DGV.Columns(2).Width = 30
        Frm_Cat_Condecoraciones.DGV.Columns(2).Visible = True
        Frm_Cat_Condecoraciones.DGV.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Cat_Condecoraciones.DGV.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Frm_Cat_Condecoraciones.DGV.Columns(2).DefaultCellStyle.BackColor = Color.Black

        Frm_Cat_Condecoraciones.DGV.Columns(3).HeaderText = "Codigo"
        Frm_Cat_Condecoraciones.DGV.Columns(3).Width = 100
        Frm_Cat_Condecoraciones.DGV.Columns(3).Visible = True
        Frm_Cat_Condecoraciones.DGV.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Cat_Condecoraciones.DGV.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Cat_Condecoraciones.DGV.Columns(3).DefaultCellStyle.BackColor = Color.Khaki

        Frm_Cat_Condecoraciones.DGV.Columns(4).HeaderText = "Descripcion"
        Frm_Cat_Condecoraciones.DGV.Columns(4).Width = 700
        Frm_Cat_Condecoraciones.DGV.Columns(4).Visible = True
        Frm_Cat_Condecoraciones.DGV.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Cat_Condecoraciones.DGV.Columns(4).DefaultCellStyle.BackColor = Color.LightGreen
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        Dim I As Integer
        For I = 0 To Frm_Cat_Condecoraciones.DGV.RowCount - 1
            If Frm_Cat_Condecoraciones.DGV.Item(2, I).Value = c32ID_CONDE Then
                Frm_Cat_Condecoraciones.DGV.Rows(I).Selected = True
                Frm_Cat_Condecoraciones.DGV.CurrentCell = Frm_Cat_Condecoraciones.DGV.Item(2, I)
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
            MiSqlCommand.CommandText = "INSERT INTO [CAT DE CONDECORACIONES] (ID_C_CONDE, ID_CONDE, CODIGO, DESCRIPCION)" &
                                  "values (" & Me.ComboBox2.SelectedValue & ", " & CInt(Me.TextBox1.Text) & ", '" & Trim(Me.TextBox3.Text) & "', '" & Trim(Me.TextBox4.Text) & "')"
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