Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Expedientes_Activos_Buscar
    Private Sub Frm_Expedientes_Activos_Buscar_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Expedientes_Activos_Buscar_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        RemoveHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
        'RemoveHandler Me.ComboBox2.SelectedIndexChanged, AddressOf ComboBox2_SelectedIndexChanged
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_1()
        Me.ComboBox1.Text = Frm_Expedientes_Activos.TextBox29.Text
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_2()
        Me.ComboBox2.Text = Frm_Expedientes_Activos.TextBox28.Text
        Call CARGAR_COMBO_VISTA_MAESTRO_DE_CARGOS_SIT()
        Me.ComboBox3.Text = Frm_Expedientes_Activos.TextBox43.Text
        AddHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
        'AddHandler Me.ComboBox2.SelectedIndexChanged, AddressOf ComboBox2_SelectedIndexChanged
        Me.TextBox1.Text = Frm_Expedientes_Activos.TextBox44.Text
        ComboBox1.Focus()
    End Sub
    Private Sub CARGAR_COMBO_VISTA_MAESTRO_DE_CARGOS_SIT()
        If Conexion.CxSPYC.State = ConnectionState.Open Then
            Conexion.CBD_SPYC()
        End If
        Conexion.ABD_SPYC()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            'MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [VISTA MAESTRO DE CARGOS SIT] where ID_N1 = " & Me.ComboBox1.SelectedValue & " and ID_N2 = " & Me.ComboBox2.SelectedValue & " order by ORDEN", Conexion.CxSPYC)
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE CARGOS DE ESTRUCTURA] order by DESCRIPCION", Conexion.CxSPYC)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox3
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_CARGO_ES"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxSPYC.State = ConnectionState.Open Then
                Conexion.CBD_SPYC()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_1()
        If Conexion.CxSPYC.State = ConnectionState.Open Then
            Conexion.CBD_SPYC()
        End If
        Conexion.ABD_SPYC()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTRUCTURA NIVEL 1] order by ORDEN", Conexion.CxSPYC)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_N1"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxSPYC.State = ConnectionState.Open Then
                Conexion.CBD_SPYC()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_2()
        If Conexion.CxSPYC.State = ConnectionState.Open Then
            Conexion.CBD_SPYC()
        End If
        Conexion.ABD_SPYC()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            'MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTRUCTURA NIVEL 2] order by ORDEN", Conexion.CxSPYC)
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTRUCTURA NIVEL 2] WHERE ID_N1 = " & Me.ComboBox1.SelectedValue & " order by ORDEN", Conexion.CxSPYC)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count <> 0 Then
                With Me.ComboBox2
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_N2"
                End With
            Else
                Me.ComboBox2.Text = ""
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxSPYC.State = ConnectionState.Open Then
                Conexion.CBD_SPYC()
            End If
        End Try
    End Sub
    Private Sub Button7_Click(sender As Object, e As System.EventArgs) Handles Button7.Click
        Me.Close()
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_2()
    End Sub
    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar correctamente el Nivel 1", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox2.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar correctamente el Nivel 2", vbInformation, "Mensaje del Sistema")
            Me.ComboBox2.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox3.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar correctamente el Cargo", vbInformation, "Mensaje del Sistema")
            Me.ComboBox3.Focus()
            Exit Sub
        End If
        Frm_Expedientes_Activos.TextBox29.Text = Me.ComboBox1.Text
        Frm_Expedientes_Activos.TextBox28.Text = Me.ComboBox2.Text
        Frm_Expedientes_Activos.TextBox43.Text = Me.ComboBox3.Text
        Frm_Expedientes_Activos.TextBox44.Text = Trim(Me.TextBox1.Text)
        Me.Close()
    End Sub
    Private Sub ComboBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox1_DoubleClick(sender As Object, e As EventArgs) Handles TextBox1.DoubleClick
        Frm_Buscar_en_RRHH.ShowDialog()
        If DATO_EN_RRHH = True Then
            Me.TextBox1.Text = rbEMPLEADO_A
        End If
    End Sub
End Class