Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Dosimetria_Upd
    Private Sub Frm_Dosimetria_Upd_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox8.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Dosimetria_Upd_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox8.Text = h5DEPARTAMENTO
        Me.TextBox9.Text = h5SERVICIO
        Me.TextBox10.Text = h5CARGO
        Me.TextBox1.Text = h5ID_DOS
        Me.TextBox2.Text = h5CODIGO_DOSIMETRO
        Me.DateTimePicker1.Value = Mid(h5FECHA_LECTURA, 1, 10)
        Me.DateTimePicker10.Value = Mid(h5FECHA_PE_I, 1, 10)
        Me.DateTimePicker20.Value = Mid(h5FECHA_PE_F, 1, 10)
        Me.DateTimePicker2.Value = Mid(h5FECHA_C_PROGRAMADA, 1, 10)
        Me.TextBox3.Text = h5DOSIS_EP
        Me.TextBox5.Text = h5DOSIS_APS
        Me.TextBox7.Text = h5DOSIS_ASS
        Me.TextBox6.Text = h5DOSIS_AA
        Me.DateTimePicker3.Value = Mid(h5FECHA_I_PORTADOR, 1, 10)
        Call CARGAR_COMBO_CAT_DE_PERIODO_DOSIMETRO()
        Me.ComboBox1.Text = h5D_PER_DOSIME
        Me.TextBox4.Text = h5OBSERVACION
        Me.TextBox8.Focus()
    End Sub
    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        Me.TextBox8.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_PERIODO_DOSIMETRO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE PERIODO DOSIMETRO] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_PER_DOSIME"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "0" Or TextBox1.Text = "" Then
            MsgBox("No se ha generado el Codigo del registro", vbInformation, "Mensaje del Sistema")
            Me.TextBox1.Text = "0"
            Me.Button2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox2.Text) = 0 Then
            Me.TextBox2.Text = "0"
        End If
        If Len(Me.TextBox3.Text) = 0 Then
            Me.TextBox3.Text = "0"
        End If
        If Len(Me.TextBox5.Text) = 0 Then
            Me.TextBox5.Text = "0"
        End If
        If Len(Me.TextBox7.Text) = 0 Then
            Me.TextBox7.Text = "0"
        End If
        If Len(Me.TextBox6.Text) = 0 Then
            Me.TextBox6.Text = "0"
        End If
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Periodo Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox4.Text) = 0 Then
            Me.TextBox4.Text = "."
        End If
        Call EDITAR_REGISTRO()
        h5ID_DOS = Me.TextBox1.Text
        CERRAR = False
        Me.Close()
    End Sub
    Private Sub EDITAR_REGISTRO()
        Dim fFECHA_1 As String
        fFECHA_1 = Mid(Me.DateTimePicker1.Value, 1, 10)
        Dim f1 As Object = If(fFECHA_1 <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_1 + "', 105), 23)", "NULL")
        Dim fFECHA_2 As String
        fFECHA_2 = Mid(Me.DateTimePicker10.Value, 1, 10)
        Dim f2 As Object = If(fFECHA_2 <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_2 + "', 105), 23)", "NULL")
        Dim fFECHA_3 As String
        fFECHA_3 = Mid(Me.DateTimePicker20.Value, 1, 10)
        Dim f3 As Object = If(fFECHA_3 <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_3 + "', 105), 23)", "NULL")
        Dim fFECHA_4 As String
        fFECHA_4 = Mid(Me.DateTimePicker2.Value, 1, 10)
        Dim f4 As Object = If(fFECHA_4 <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_4 + "', 105), 23)", "NULL")
        Dim fFECHA_5 As String
        fFECHA_5 = Mid(Me.DateTimePicker3.Value, 1, 10)
        Dim f5 As Object = If(fFECHA_5 <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_5 + "', 105), 23)", "NULL")

        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE HIGIENE DOSIMETRIA] SET DEPARTAMENTO = '" & Me.TextBox8.Text & "', SERVICIO = '" & Me.TextBox9.Text & "', CARGO = '" & Me.TextBox10.Text & "', ID_M_P = " & Frm_Dosimetria.TextBox36.Text & ", CODIGO_DOSIMETRO = '" & Me.TextBox2.Text & "', FECHA_LECTURA = " & f1 & ", FECHA_PE_I = " & f2 & ", FECHA_PE_F = " & f3 & ", DOSIS_EP = '" & Me.TextBox3.Text & "', DOSIS_APS = '" & Me.TextBox5.Text & "', DOSIS_ASS = '" & Me.TextBox7.Text & "', DOSIS_AA = '" & Me.TextBox6.Text & "', FECHA_I_PORT = " & f5 & ", FECHA_C_PROGRAMADA = " & f4 & ", ID_PER_DOSIME = " & Me.ComboBox1.SelectedValue & ", OBSERVACION = '" & Me.TextBox4.Text & "' WHERE ID_DOS = " & CInt(Me.TextBox1.Text) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
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
    Private Sub DateTimePicker10_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker10.KeyDown
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
    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
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
    Private Sub DateTimePicker1_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub DateTimePicker20_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker20.KeyDown
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
            Me.TextBox6.Text = Val(Me.TextBox5.Text) + Val(Me.TextBox7.Text)
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub DateTimePicker3_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker3.KeyDown
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
End Class