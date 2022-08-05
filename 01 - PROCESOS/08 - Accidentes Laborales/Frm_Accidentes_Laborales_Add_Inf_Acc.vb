Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Accidentes_Laborales_Add_Inf_Acc
    Private Sub Frm_Accidentes_Laborales_Add_Inf_Acc_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox3.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Accidentes_Laborales_Add_Inf_Acc_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = Frm_Accidentes_Laborales.TextBox31.Text
        Me.TextBox6.Text = ""
        Me.TextBox7.Text = Frm_Accidentes_Laborales.TextBox11.Text
        Me.TextBox8.Text = Frm_Accidentes_Laborales.TextBox12.Text
        Me.TextBox9.Text = Frm_Accidentes_Laborales.TextBox13.Text
        Me.TextBox10.Text = ""
        Me.TextBox11.Text = ""
        Me.TextBox3.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.TextBox3.Focus()
        CERRAR = True
        Me.Close()
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
    Private Sub GENERAR_CODIGO()
        CODIGO = 0
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            MiDataAdapter = New SqlDataAdapter("SELECT ID_M_IT FROM [MAESTRO DE HIGIENE INFORMACION DEL ACCIDENTE] ORDER BY ID_M_IT", Conexion.CxRRHH)
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
        Call GRABAR_REGISTRO()
        h2ID_M_IT = Me.TextBox1.Text
        CERRAR = false
        Me.Close()
    End Sub
    Private Sub GRABAR_REGISTRO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE HIGIENE INFORMACION DEL ACCIDENTE] (ID_M_IT, ID_M_P, NO_PARTE_ACC, NO_REP_NSS, NO_REP_MITRAB, EDAD, ACCIDENTE_ANTERIOR, DEPARTAMENTO, SERVICIO, CARGO, ANT_EMPRESA_MESES, ANT_PUESTO_MESES)" &
                                  "values (" & CInt(Me.TextBox1.Text) & ", " & Frm_Accidentes_Laborales.TextBox36.Text & ", '" & Me.TextBox3.Text & "', '" & Me.TextBox2.Text & "', '" & Me.TextBox4.Text & "', " & CInt(Me.TextBox5.Text) & ", '" & Me.TextBox6.Text & "', '" & Me.TextBox7.Text & "', '" & Me.TextBox8.Text & "', '" & Me.TextBox9.Text & "', " & Me.TextBox10.Text & ", " & Me.TextBox11.Text & ")"
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