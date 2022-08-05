Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Cat_Agregar_Estructura_N1
    Private Sub Frm_Cat_Agregar_Estructura_N1_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = ""
            Me.TextBox2.Text = ""
            Me.TextBox3.Text = ""
            Me.TextBox4.Text = ""
            Me.TextBox2.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Cat_Agregar_Estructura_N1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = ""
        Call GENERAR_ORDEN()
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_ESTRUCTURA()
        Me.TextBox2.Focus()
        Me.TextBox2.SelectAll()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_ESTRUCTURA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE ESTRUCTURA] order by ID_T_ES", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_T_ES"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox2.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
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
    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub GENERAR_ORDEN()
        CODIGO = 0
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            MiDataAdapter = New SqlDataAdapter("SELECT ORDEN FROM [CAT DE ESTRUCTURA NIVEL 1] ORDER BY ORDEN", Conexion.CxRRHH)
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
            Me.TextBox2.Text = CODIGO
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
            MiDataAdapter = New SqlDataAdapter("SELECT ID_N1 FROM [CAT DE ESTRUCTURA NIVEL 1] ORDER BY ID_N1", Conexion.CxRRHH)
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
        If Len(Me.TextBox2.Text) = 0 Then
            MsgBox("Debe digitar el No. de Orden Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox3.Text) = 0 Then
            MsgBox("Debe digitar la Descripción Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox3.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox4.Text) = 0 Then
            MsgBox("Debe digitar la Sigla Correctamente o escribir (.) si no desea una Sigla para este registro", vbInformation, "Mensaje del Sistema")
            Me.TextBox4.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Tipo Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        Call GRABAR_REGISTRO()
        Call REENUMERAR
        c2ID_N1 = Me.TextBox1.Text
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
        Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        Me.TextBox1.Text = ""
        Call GENERAR_ORDEN()
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox2.Focus()
        Me.TextBox2.SelectAll()
        CERRAR = False
        'Me.Close()
    End Sub
    Dim iINDICE As Integer
    Dim iNORDEN As String
    Private Sub REENUMERAR()
        Dim CONTADOR As Integer
        CONTADOR = 0
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT ID_N1, ORDEN FROM [dbo].[CAT DE ESTRUCTURA NIVEL 1] ORDER BY [ORDEN]", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                For I = 0 To MiDataTable.Rows.Count - 1
                    CONTADOR = CONTADOR + 1
                    iINDICE = MiDataTable.Rows(I).Item(0).ToString
                    iNORDEN = CONTADOR
                    Call ACTUALIZAR_CAT_DE_ESTRUCTURA_NIVEL_1()
                Next
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ACTUALIZAR_CAT_DE_ESTRUCTURA_NIVEL_1()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [CAT DE ESTRUCTURA NIVEL 1] SET ORDEN = '" & iNORDEN & "' WHERE ID_N1 = " & CInt(iINDICE) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub GRABAR_REGISTRO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [CAT DE ESTRUCTURA NIVEL 1] (ID_N1, ORDEN, DESCRIPCION, SIGLAS, ID_T_ES, MIXTA, MILITAR, PAME)" &
                                  "values (" & CInt(Me.TextBox1.Text) & ", '" & Trim(Me.TextBox2.Text) & "', '" & Trim(Me.TextBox3.Text) & "', '" & Trim(Me.TextBox4.Text) & "', " & Me.ComboBox1.SelectedValue & ", '" & Me.CheckBox1.Checked & "', '" & Me.CheckBox2.Checked & "', '" & Me.CheckBox3.Checked & "')"
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
        Dim MiAdaptador As New SqlClient.SqlDataAdapter("Select * from [VISTA CAT DE ESTRUCTURA NIVEL 1] ORDER BY ORDEN", Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA CAT DE ESTRUCTURA NIVEL 1]")
        Frm_Cat_Estructura_N1.DGV.DataSource = MiDataSet.Tables("[VISTA CAT DE ESTRUCTURA NIVEL 1]")
        Frm_Cat_Estructura_N1.Label2.Text = "Total de Registros: " & Frm_Cat_Estructura_N1.DGV.RowCount
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW()
        Frm_Cat_Estructura_N1.DGV.Columns(0).HeaderText = "Id"
        Frm_Cat_Estructura_N1.DGV.Columns(0).Width = 30
        Frm_Cat_Estructura_N1.DGV.Columns(0).Visible = True
        Frm_Cat_Estructura_N1.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Cat_Estructura_N1.DGV.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Frm_Cat_Estructura_N1.DGV.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Frm_Cat_Estructura_N1.DGV.Columns(1).HeaderText = "Orden"
        Frm_Cat_Estructura_N1.DGV.Columns(1).Width = 40
        Frm_Cat_Estructura_N1.DGV.Columns(1).Visible = True
        Frm_Cat_Estructura_N1.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Cat_Estructura_N1.DGV.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Frm_Cat_Estructura_N1.DGV.Columns(1).DefaultCellStyle.BackColor = Color.Khaki

        Frm_Cat_Estructura_N1.DGV.Columns(2).HeaderText = "Descripcion"
        Frm_Cat_Estructura_N1.DGV.Columns(2).Width = 450
        Frm_Cat_Estructura_N1.DGV.Columns(2).Visible = True
        Frm_Cat_Estructura_N1.DGV.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Cat_Estructura_N1.DGV.Columns(2).DefaultCellStyle.BackColor = Color.LightGreen
        'Frm_Cat_Estructura_N1.DGV.Columns(2).Frozen = True

        Frm_Cat_Estructura_N1.DGV.Columns(3).HeaderText = "Siglas"
        Frm_Cat_Estructura_N1.DGV.Columns(3).Width = 100
        Frm_Cat_Estructura_N1.DGV.Columns(3).Visible = True
        Frm_Cat_Estructura_N1.DGV.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Cat_Estructura_N1.DGV.Columns(4).HeaderText = "ID_T_ES"
        Frm_Cat_Estructura_N1.DGV.Columns(4).Width = 0
        Frm_Cat_Estructura_N1.DGV.Columns(4).Visible = False
        Frm_Cat_Estructura_N1.DGV.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Cat_Estructura_N1.DGV.Columns(5).HeaderText = "Tipo"
        Frm_Cat_Estructura_N1.DGV.Columns(5).Width = 200
        Frm_Cat_Estructura_N1.DGV.Columns(5).Visible = True
        Frm_Cat_Estructura_N1.DGV.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Cat_Estructura_N1.DGV.Columns(6).HeaderText = "Mixta"
        Frm_Cat_Estructura_N1.DGV.Columns(6).Width = 70
        Frm_Cat_Estructura_N1.DGV.Columns(6).Visible = True
        Frm_Cat_Estructura_N1.DGV.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Cat_Estructura_N1.DGV.Columns(7).HeaderText = "Militar"
        Frm_Cat_Estructura_N1.DGV.Columns(7).Width = 70
        Frm_Cat_Estructura_N1.DGV.Columns(7).Visible = True
        Frm_Cat_Estructura_N1.DGV.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Cat_Estructura_N1.DGV.Columns(8).HeaderText = "Pame"
        Frm_Cat_Estructura_N1.DGV.Columns(8).Width = 70
        Frm_Cat_Estructura_N1.DGV.Columns(8).Visible = True
        Frm_Cat_Estructura_N1.DGV.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        Dim I As Integer
        For I = 0 To Frm_Cat_Estructura_N1.DGV.RowCount - 1
            If Frm_Cat_Estructura_N1.DGV.Item(0, I).Value = c2ID_N1 Then
                Frm_Cat_Estructura_N1.DGV.Rows(I).Selected = True
                Frm_Cat_Estructura_N1.DGV.CurrentCell = Frm_Cat_Estructura_N1.DGV.Item(0, I)
                Exit Sub
            End If
        Next
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
End Class