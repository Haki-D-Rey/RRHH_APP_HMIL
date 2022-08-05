Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Accidentes_Laborales_Upd_Inf_Acc
    Private Sub Frm_Accidentes_Laborales_Upd_Inf_Acc_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = h2ID_M_IT
        Me.TextBox3.Text = h2NO_PARTE_ACC
        Me.TextBox2.Text = h2NO_REP_NSS
        Me.TextBox4.Text = h2NO_REP_MITRAB
        Me.TextBox5.Text = h2EDAD
        Me.TextBox6.Text = h2ACCIDENTE_ANTERIOR
        Me.TextBox7.Text = h2DEPARTAMENTO
        Me.TextBox8.Text = h2SERVICIO
        Me.TextBox9.Text = h2CARGO
        Me.TextBox10.Text = h2ANT_EMPRESA_MESES
        Me.TextBox11.Text = h2ANT_PUESTO_MESES
        Me.TextBox3.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox3.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub Frm_Accidentes_Laborales_Upd_Inf_Acc_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox3.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
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
    Private Sub TextBox11_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox11.KeyDown
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
        If Len(Me.TextBox3.Text) = 0 Then
            MsgBox("Debe Digitar el valor del No. de Parte de Accidente", vbInformation, "Mensaje del Sistema")
            Me.TextBox3.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox2.Text) = 0 Then
            MsgBox("Debe Digitar el valor del No. de Reporte INSS", vbInformation, "Mensaje del Sistema")
            Me.TextBox2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox4.Text) = 0 Then
            MsgBox("Debe Digitar el valor del No. de Reporte MITRAB", vbInformation, "Mensaje del Sistema")
            Me.TextBox4.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox5.Text) = 0 Then
            MsgBox("Debe Digitar el valor de la Edad", vbInformation, "Mensaje del Sistema")
            Me.TextBox5.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox6.Text) = 0 Then
            Me.TextBox6.Text = "."
        End If
        If Len(Me.TextBox7.Text) = 0 Then
            MsgBox("Debe Digitar el Departamento", vbInformation, "Mensaje del Sistema")
            Me.TextBox7.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox8.Text) = 0 Then
            MsgBox("Debe Digitar el Servicio", vbInformation, "Mensaje del Sistema")
            Me.TextBox8.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox9.Text) = 0 Then
            MsgBox("Debe Digitar el Cargo", vbInformation, "Mensaje del Sistema")
            Me.TextBox9.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox10.Text) = 0 Then
            MsgBox("Debe Digitar el Valor de la Antiguedad en la Empresa", vbInformation, "Mensaje del Sistema")
            Me.TextBox10.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox11.Text) = 0 Then
            MsgBox("Debe Digitar el valor de la Antiguedad en el Cargo", vbInformation, "Mensaje del Sistema")
            Me.TextBox11.Focus()
            Exit Sub
        End If
        Call EDITAR_REGISTRO()
        h2ID_M_IT = Me.TextBox1.Text
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
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE HIGIENE INFORMACION DEL ACCIDENTE] SET NO_PARTE_ACC = '" & Me.TextBox3.Text & "', NO_REP_NSS = '" & Me.TextBox2.Text & "', NO_REP_MITRAB = '" & Me.TextBox4.Text & "', EDAD = " & CInt(Me.TextBox5.Text) & ", ACCIDENTE_ANTERIOR = '" & Me.TextBox6.Text & "', DEPARTAMENTO = '" & Me.TextBox7.Text & "', SERVICIO = '" & Me.TextBox8.Text & "', CARGO = '" & Me.TextBox9.Text & "', ANT_EMPRESA_MESES = " & CInt(Me.TextBox10.Text) & ", ANT_PUESTO_MESES = " & CInt(Me.TextBox11.Text) & " WHERE ID_M_IT = " & Me.TextBox1.Text & ""
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