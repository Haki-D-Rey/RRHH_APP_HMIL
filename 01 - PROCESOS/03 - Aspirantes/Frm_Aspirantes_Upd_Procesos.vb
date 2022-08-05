Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Aspirantes_Upd_Procesos
    Private Sub Frm_Aspirantes_Upd_Procesos_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.DateTimePicker1.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Aspirantes_Upd_Procesos_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        RemoveHandler ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
        Me.DateTimePicker1.Value = asppFECHA_P
        Call CARGAR_COMBO_SOL_CAT_DE_PROCESOS_PRINCIPALES()
        Me.ComboBox1.Text = asppD_PROC
        Call CARGAR_COMBO_SOL_CAT_DE_UBICACION_DE_RECURSOS()
        Me.ComboBox2.Text = asppD_U_RECURSO
        Me.CheckBox1.Checked = asppREALIZADA
        Call CARGAR_COMBO_SOL_CAT_DE_RESULTADOS()
        Me.ComboBox3.Text = asppD_R_SOL
        Me.TextBox2.Text = asppOBSERVACIONES
        Me.TextBox1.Text = asppID_S_P
        Me.DateTimePicker1.Focus()
        AddHandler ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.DateTimePicker1.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_SOL_CAT_DE_PROCESOS_PRINCIPALES()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [SOL_CAT DE PROCESOS PRINCIPALES] ORDER BY ID_PROC", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_PROC"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_SOL_CAT_DE_UBICACION_DE_RECURSOS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [SOL_CAT DE UBICACION DE RECURSOS] ORDER BY ID_U_RECURSO", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox2
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_U_RECURSO"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_SOL_CAT_DE_RESULTADOS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [SOL_CAT DE RESULTADOS] WHERE ID_PROC = " & Me.ComboBox1.SelectedValue & " ORDER BY ID_R_SOL", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox3
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_R_SOL"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Call CARGAR_COMBO_SOL_CAT_DE_RESULTADOS()
    End Sub
    Private Sub DateTimePicker1_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker1.KeyDown
        If e.KeyCode = Keys.Enter Then
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
    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub CheckBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "0" Or TextBox1.Text = "" Then
            MsgBox("No se ha generado el Codigo del registro", vbInformation, "Mensaje del Sistema")
            Me.TextBox1.Text = "0"
            Me.Button2.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Proceso actual en el que se ecuentra el recurso", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox2.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar la Ubicación actual en la que se ecuentra el recurso", vbInformation, "Mensaje del Sistema")
            Me.ComboBox2.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox3.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Resultado del Proceso en el que se ecuentra el recurso", vbInformation, "Mensaje del Sistema")
            Me.ComboBox3.Focus()
            Exit Sub
        End If
        Call Me.EDITAR_REGISTRO()
        asppID_S_P = Me.TextBox1.Text
        If Me.ComboBox1.Text = "RECEPCION DE DOCUMENTOS" Then
            Call EDITAR_PRINCIPAL()
        End If
        Me.DateTimePicker1.Focus()
        CERRAR = False
        Me.Close()
    End Sub
    Private Sub EDITAR_PRINCIPAL()
        Dim fFEC_ING As String = Mid(Me.DateTimePicker1.Value, 1, 10)
        If Me.DateTimePicker1.Checked = True Then
            fFEC_ING = Mid(Me.DateTimePicker1.Value, 1, 10)
        Else
            fFEC_ING = "NULL"
        End If
        Dim fechaingreso As Object = If(fFEC_ING <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFEC_ING + "', 105), 23)", "NULL")

        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [SOL_MAESTRO DE SOLICITANTES] SET FECHA_I = " & fechaingreso & " WHERE ID_S = " & aspID_S & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub EDITAR_REGISTRO()
        Dim FECHA01 As String = Mid(Me.DateTimePicker1.Value, 1, 10)
        Dim f1 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA01 + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [SOL_MAESTRO DE SOLICITANTES PROCESOS] SET FECHA_P = " & f1 & ", ID_PROC = " & Me.ComboBox1.SelectedValue & ", ID_U_RECURSO = " & Me.ComboBox2.SelectedValue & ", REALIZADA = '" & Me.CheckBox1.Checked & "', ID_R_SOL = " & Me.ComboBox3.SelectedValue & ", OBSERVACIONES = '" & Me.TextBox2.Text & "' WHERE ID_S_P = " & Me.TextBox1.Text & ""
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