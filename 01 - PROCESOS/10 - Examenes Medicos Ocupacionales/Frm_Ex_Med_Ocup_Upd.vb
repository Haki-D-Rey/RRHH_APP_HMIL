Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Ex_Med_Ocup_Upd
    Private Sub Frm_Ex_Med_Ocup_Upd_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            CERRAR = True
            Me.DateTimePicker1.Focus()
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Ex_Med_Ocup_Upd_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = emoID_EMO
        Me.DateTimePicker1.Value = Mid(emoFECHA_EVA, 1, 10)
        Me.TextBox2.Text = emoNO_EXPEDIENTE
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_EVALUACION()
        Me.ComboBox1.Text = emoD_T_EVALUACION
        Me.TextBox3.Text = emoCARGO_ANTERIOR
        Me.TextBox4.Text = emoCARGO_ACTUAL
        Me.TextBox6.Text = emoAE_ANTERIOR
        Me.TextBox5.Text = emoAE_ACTUAL
        Me.TextBox7.Text = emoDIAGNOSTICO
        Me.TextBox8.Text = emoRECOMENDACIONES
        Call CARGAR_COMBO_CAT_DE_RESULTADO_DE_EVALUACION()
        Me.ComboBox2.Text = emoD_R_EVALUACION
        Me.TextBox9.Text = emoOBSERVACIONES
        Me.DateTimePicker1.Focus()
    End Sub
    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        CERRAR = True
        Me.DateTimePicker1.Focus()
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_RESULTADO_DE_EVALUACION()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE RESULTADO DE EVALUACION] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox2
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_R_EVALUACION"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_EVALUACION()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE EVALUACION] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_T_EVALUACION"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub DateTimePicker1_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker1.KeyDown
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
    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
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
    Private Sub TextBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox6.KeyDown
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
    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "0" Or TextBox1.Text = "" Then
            MsgBox("No se ha generado el Codigo del registro", vbInformation, "Mensaje del Sistema")
            Me.TextBox1.Text = "0"
            Me.Button2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox2.Text) = 0 Then
            MsgBox("Debe Digitar el No. de Expediente", vbInformation, "Mensaje del Sistema")
            Me.TextBox2.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Tipo de Evaluación Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox3.Text) = 0 Then
            MsgBox("Debe Digitar el Cargo Anterior Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox3.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox4.Text) = 0 Then
            MsgBox("Debe Digitar el Cargo Actual Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox4.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox6.Text) = 0 Then
            MsgBox("Debe Digitar los Años de Exposición Anterior Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox6.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox5.Text) = 0 Then
            MsgBox("Debe Digitar los Años de Exposición Actual Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox5.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox7.Text) = 0 Then
            MsgBox("Debe Digitar el Diagnostico Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox7.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox8.Text) = 0 Then
            MsgBox("Debe Digitar las Recomendaciones Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox8.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox2.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Resultado de la Evaluacion Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox9.Text) = 0 Then
            MsgBox("Debe Digitar las Observaciones Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox9.Focus()
            Exit Sub
        End If

        Call EDITAR_REGISTRO()
        emoID_EMO = Me.TextBox1.Text
        Me.TextBox1.Text = "0"
        Me.TextBox2.Text = ""
        'Call CARGAR_COMBO_CAT_DE_TIPO_DE_EVALUACION()
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox6.Text = ""
        Me.TextBox5.Text = ""
        Me.TextBox7.Text = ""
        Me.TextBox8.Text = ""
        'Call CARGAR_COMBO_CAT_DE_RESULTADO_DE_EVALUACION()
        Me.TextBox9.Text = ""
        Me.DateTimePicker1.Focus()
        CERRAR = False
        Me.Close()
    End Sub
    Private Sub EDITAR_REGISTRO()
        Dim fFECHA_1 As String
        fFECHA_1 = Mid(Me.DateTimePicker1.Value, 1, 10)
        Dim f1 As Object = If(fFECHA_1 <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_1 + "', 105), 23)", "NULL")

        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE HIGIENE EXAMEN MEDICO OCUPACIONAL] SET DEPARTAMENTO = '" & Frm_Ex_Med_Ocup.TextBox6.Text & "', SERVICIO = '" & Frm_Ex_Med_Ocup.TextBox7.Text & "', CARGO = '" & Frm_Ex_Med_Ocup.TextBox4.Text & "', FECHA_EVA = " & f1 & ", NO_EXPEDIENTE = '" & Me.TextBox2.Text & "', ID_T_EVALUACION = " & Me.ComboBox1.SelectedValue & ", CARGO_ANTERIOR = '" & Me.TextBox3.Text & "', CARGO_ACTUAL = '" & Me.TextBox4.Text & "', AE_ANTERIOR = '" & Me.TextBox6.Text & "', AE_ACTUAL = '" & Me.TextBox5.Text & "', DIAGNOSTICO = '" & Me.TextBox7.Text & "', RECOMENDACIONES = '" & Me.TextBox8.Text & "', ID_R_EVALUACION = " & Me.ComboBox1.SelectedValue & ", OBSERVACIONES = '" & Me.TextBox9.Text & "' WHERE ID_EMO = " & CInt(Me.TextBox1.Text) & ""
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