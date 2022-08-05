Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Editar_Experiencia_Laboral
    Private Sub Frm_Editar_Experiencia_Laboral_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = ""
            Me.TextBox2.Text = ""
            Me.TextBox3.Text = ""
            Me.TextBox4.Text = "0"
            Me.TextBox5.Text = ""
            Me.TextBox6.Text = ""
            Me.TextBox7.Text = ""
            Me.DateTimePicker1.Value = Now
            Me.DateTimePicker2.Value = Now
            Me.TextBox8.Text = "0"
            Me.TextBox9.Text = ""
            Me.TextBox10.Text = ""
            Me.TextBox2.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Editar_Experiencia_Laboral_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = vmelID_EL
        Me.TextBox2.Text = vmelEMPRESA_NOMBRE
        Me.TextBox3.Text = vmelEMPRESA_DIRECCION
        Me.TextBox4.Text = vmelEMPRESA_TELEFONO
        Me.TextBox5.Text = vmelEMPRESA_ACTIVIDAD
        Me.TextBox6.Text = vmelCARGO_DESEM
        Me.TextBox7.Text = vmelJEFE_INMEDIATO
        Me.DateTimePicker1.Value = vmelFEC_INI
        Me.DateTimePicker2.Value = vmelFEC_FIN
        Me.TextBox8.Text = vmelSALARIO
        Me.TextBox9.Text = vmelCAUSA_RETIRO
        Me.TextBox10.Text = vmelOBLI_RESP
        Me.TextBox2.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = "0"
        Me.TextBox5.Text = ""
        Me.TextBox6.Text = ""
        Me.TextBox7.Text = ""
        Me.DateTimePicker1.Value = Now
        Me.DateTimePicker2.Value = Now
        Me.TextBox8.Text = "0"
        Me.TextBox9.Text = ""
        Me.TextBox10.Text = ""
        Me.TextBox2.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox6.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox7.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub DateTimePicker1_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub DateTimePicker2_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox9.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox10.KeyDown
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
        If Len(Me.TextBox2.Text) = 0 Then
            MsgBox("Debe digitar el Nombre de la Empresa correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox3.Text) = 0 Then
            MsgBox("Debe digitar la Dirección de la Empresa correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox3.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox4.Text) = 0 Then
            Me.TextBox4.Text = ""
        End If
        If Len(Me.TextBox5.Text) = 0 Then
            MsgBox("Debe digitar la Actividad de la Empresa correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox5.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox6.Text) = 0 Then
            MsgBox("Debe digitar el Cargo Desempeñado en la Empresa correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox6.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox7.Text) = 0 Then
            MsgBox("Debe digitar el Nombre del Jefe Inmediato correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox7.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox8.Text) = 0 Or Me.TextBox8.Text = "0" Then
            MsgBox("Debe digitar el Salario devengado correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox8.Focus()
        End If
        If Len(Me.TextBox9.Text) = 0 Then
            MsgBox("Debe digitar los Motivos de retiro correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox9.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox10.Text) = 0 Then
            MsgBox("Debe digitar las Obligaciones y Responsabilidades correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox10.Focus()
            Exit Sub
        End If
        Call EDITAR_REGISTRO()
        vmnfID_NF = Me.TextBox1.Text
        CERRAR = False
        Me.Close()
    End Sub
    Private Sub EDITAR_REGISTRO()
        Dim fFEC_INI As String = Me.DateTimePicker1.Text
        Dim fFEC_FIN As String = Me.DateTimePicker2.Text
        Dim fechaInicial = "Convert(VARCHAR(10), Convert(Date, '" + fFEC_INI + "', 105), 23)"
        Dim fechaFinal = "Convert(VARCHAR(10), Convert(Date, '" + fFEC_FIN + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE EXPERIENCIA LABORAL] SET EMPRESA_NOMBRE = '" & Me.TextBox2.Text & "', EMPRESA_DIRECCION = '" & Me.TextBox3.Text & "', EMPRESA_TELEFONO = '" & Me.TextBox4.Text & "', EMPRESA_ACTIVIDAD = '" & Me.TextBox5.Text & "', CARGO_DESEM = '" & Me.TextBox6.Text & "', JEFE_INMEDIATO = '" & Me.TextBox7.Text & "', SALARIO = '" & Me.TextBox8.Text & "', FEC_INI = " & fechaInicial & ", FEC_FIN = " & fechaFinal & ", CAUSA_RETIRO = '" & Me.TextBox9.Text & "', OBLI_RESP = '" & Me.TextBox10.Text & "' WHERE ID_EL = " & Me.TextBox1.Text & ""
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