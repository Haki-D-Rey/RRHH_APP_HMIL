Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Cat_T_R_Examen_Upd
    Private Sub Frm_Cat_T_R_Examen_Upd_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = ""
            Me.TextBox2.Text = ""
            Me.TextBox2.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Cat_T_R_Examen_Upd_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = c44ID_F
        Me.TextBox2.Text = c44DESCRIPCION
        Me.TextBox2.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "0" Or TextBox1.Text = "" Then
            MsgBox("No se ha generado el Codigo del registro", vbInformation, "Mensaje del Sistema")
            Me.TextBox1.Text = "0"
            Me.Button2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox2.Text) = 0 Then
            MsgBox("Debe digitar la Descripción Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox2.Focus()
            Exit Sub
        End If
        Call EDITAR_REGISTRO()
        c44ID_F = Me.TextBox1.Text
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
            MiSqlCommand.CommandText = "UPDATE [CAT DE TIPO DE RESULTADO DE EXAMEN] SET DESCRIPCION = '" & Trim(Me.TextBox2.Text) & "' WHERE ID_TRE = " & CInt(Me.TextBox1.Text) & ""
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