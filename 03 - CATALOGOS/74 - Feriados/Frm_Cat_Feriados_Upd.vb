Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Cat_Feriados_Upd
    Private Sub Frm_Cat_Feriados_Upd_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.DateTimePicker1.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Cat_Feriados_Upd_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = cID_D_F
        Me.DateTimePicker1.Value = cFECHA_F
        Me.CheckBox1.Checked = cACTIVO
        Me.DateTimePicker1.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.DateTimePicker1.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub DateTimePicker1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles DateTimePicker1.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub CheckBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CheckBox1.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker1.Value, 1, 10) + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [CAT DE FERIADOS] SET FECHA_F = " & F1 & ", ACTIVO =  '" & Me.CheckBox1.Checked & "' WHERE ID_D_F = " & cID_D_F & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
        cID_D_F = cID_D_F
        CERRAR = False
        Me.Close()
    End Sub
End Class