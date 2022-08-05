Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Cat_Administradores_R_Trasladar
    Private Sub Frm_Cat_Administradores_R_Trasladar_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.ComboBox1.Focus()
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Cat_Administradores_R_Trasladar_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = c46FDESCRIPCION
        Call CARGAR_COMBO_CAT_DE_ADMINISTRADORES
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.ComboBox1.Focus()
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ADMINISTRADORES()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ADMINISTRADORES] where ACTIVO = 'TRUE' ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_ADMIN"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.TextBox1.Text = "" Then
            MsgBox("No se encuentra el Administrador de Origen", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Administrador de Destino Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Esta seguro de Trasladar los registros del Usuario " & Me.TextBox1.Text & " al Usuario " & Me.ComboBox1.Text & "?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            Me.Cursor = Cursors.WaitCursor
            Call ACTUALIZAR_REGISTRO
            MsgBox("Traslado Satisfactorio", vbInformation, "Mensaje del Sistema")
            Me.Close()
            Me.Cursor = Cursors.Default
        End If
    End Sub
    Private Sub ACTUALIZAR_REGISTRO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE PERSONAS] SET ID_ADMIN = " & Me.ComboBox1.SelectedValue & " WHERE ID_ADMIN = " & CInt(c46ID_ADMIN) & ""
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