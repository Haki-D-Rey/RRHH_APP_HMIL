Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Establecer_Clave
    Private Sub MOSTRAR_CONTRASEÑA()
        If Me.CheckBox1.Checked = True Then
            Me.TextBox3.PasswordChar = ""
        Else
            Me.TextBox3.PasswordChar = "•"
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.TextBox3.Text = ""
        Me.TextBox3.Focus()
        Me.Close()
    End Sub
    Private Sub Frm_Establecer_Clave_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TextBox3.Text = ""
        Me.TextBox3.Focus()
    End Sub
    Private Sub Frm_Establecer_Clave_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox3.Text = ""
            Me.TextBox3.Focus()
            Me.Close()
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call ACTUALIZAR_REGISTRO()
        Me.TextBox3.Text = ""
        Me.TextBox3.Focus()
        Me.Close()
    End Sub
    Private Sub ACTUALIZAR_REGISTRO()
        CADENAsql = "UPDATE [MAESTRO DE USUARIOS] SET CLAVE = '" & ENCRIPTAR.ENCRIPTAR_PALABRA(Me.TextBox3.Text) & "' WHERE ID_USUARIO = " & CInt(uID_USUARIO) & ""
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = CADENAsql
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As System.EventArgs) Handles CheckBox1.CheckedChanged
        Call MOSTRAR_CONTRASEÑA()
    End Sub
End Class