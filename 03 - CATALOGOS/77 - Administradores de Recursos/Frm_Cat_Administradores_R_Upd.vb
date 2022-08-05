Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Cat_Administradores_R_Upd
    Private Sub Frm_Cat_Administradores_R_Upd_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.CheckBox1.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Cat_Administradores_R_Upd_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = c46ID_ADMIN
        Me.TextBox2.Text = c46FDESCRIPCION
        Me.CheckBox1.Checked = c46ACTIVO
        Me.CheckBox1.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.CheckBox1.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "0" Or TextBox1.Text = "" Then
            MsgBox("No se ha generado el Codigo del registro", vbInformation, "Mensaje del Sistema")
            Me.TextBox1.Text = "0"
            Me.Button2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox2.Text) = 0 Then
            MsgBox("No se encuentra la Descripcion del Administrador de Recursos", vbInformation, "Mensaje del Sistema")
            Me.Button2.Focus()
            Exit Sub
        End If
        Call EDITAR_REGISTRO()
        c46ID_ADMIN = Me.TextBox1.Text
        CERRAR = False
        Me.Close()
    End Sub
    Private Sub EDITAR_REGISTRO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [CAT DE ADMINISTRADORES] SET ACTIVO = '" & Me.CheckBox1.Checked & "' WHERE ID_ADMIN = " & CInt(Me.TextBox1.Text) & ""
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