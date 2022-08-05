Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Editar_Subsidios
    Private Sub Frm_Editar_Subsidios_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = "0"
            Me.TextBox2.Text = ""
            Me.TextBox3.Text = ""
            Me.TextBox4.Text = "0"
            Me.TextBox5.Text = "0"
            Me.TextBox6.Text = ""
            Me.TextBox7.Text = ""
            Me.TextBox8.Text = ""
            Me.TextBox9.Text = ""
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Editar_Subsidios_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = vmssID_SUBS
        Me.TextBox2.Text = vmssNO_EXPEDIENTE
        Me.TextBox3.Text = vmssNO_BOLETA
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_SUBSIDIOS()
        Me.ComboBox1.Text = vmssD_T_SUBS
        Me.TextBox4.Text = vmssN_ORDEN
        Me.DateTimePicker1.Value = vmssFECHA_I
        Me.DateTimePicker2.Value = vmssFECHA_F
        Me.TextBox5.Text = vmssCANT_DIAS

        If vmssFECHA_PARTO <> "" Then
            Me.DateTimePicker3.Value = vmssFECHA_PARTO
            Me.DateTimePicker3.Checked = True
        Else
            Me.DateTimePicker3.Checked = False
        End If
        If vmssFECHA_PARTO_PROB <> "" Then
            Me.DateTimePicker4.Value = vmssFECHA_PARTO_PROB
            Me.DateTimePicker4.Checked = True
        Else
            Me.DateTimePicker4.Checked = False
        End If
        If vmssFECHA_ACC_LAB <> "" Then
            Me.DateTimePicker5.Value = vmssFECHA_ACC_LAB
            Me.DateTimePicker5.Checked = True
        Else
            Me.DateTimePicker5.Checked = False
        End If
        If vmssFECHA_DECLARACION <> "" Then
            Me.DateTimePicker6.Value = vmssFECHA_DECLARACION
            Me.DateTimePicker6.Checked = True
        Else
            Me.DateTimePicker6.Checked = False
        End If

        Me.TextBox6.Text = vmssMEDICO
        Me.TextBox7.Text = vmssN_MINSA
        Call CARGAR_COMBO_CAT_DE_ESPECIALIDADES_MEDICAS()
        Me.ComboBox2.Text = vmssD_E_MED
        Me.TextBox8.Text = vmssDIAGNOSTICO
        Me.TextBox9.Text = vmssOBSERVACIONES
        Me.TextBox2.Focus()
    End Sub
    Private Sub Button7_Click(sender As Object, e As System.EventArgs) Handles Button7.Click
        Me.TextBox1.Text = "0"
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = "0"
        Me.TextBox5.Text = "0"
        Me.TextBox6.Text = ""
        Me.TextBox7.Text = ""
        Me.TextBox8.Text = ""
        Me.TextBox9.Text = ""
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESPECIALIDADES_MEDICAS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESPECIALIDADES MEDICAS] WHERE ID_T_ESTUDIO = 15  order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox2
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_E_MED"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_SUBSIDIOS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE SUBSIDIOS] order by ID_T_SUBS", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_T_SUBS"
            End With

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
    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub DateTimePicker3_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub DateTimePicker4_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker4.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub DateTimePicker5_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker5.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub DateTimePicker6_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker6.KeyDown
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
    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox3_KeyDown(sender As Object, e As KeyEventArgs)
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "0" Or TextBox1.Text = "" Then
            MsgBox("No se ha generado el Codigo del registro", vbInformation, "Mensaje del Sistema")
            Me.TextBox1.Text = "0"
            Me.Button7.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox2.Text) = 0 Then
            Me.TextBox2.Text = "0"
        End If
        If Len(Me.TextBox3.Text) = 0 Then
            Me.TextBox3.Text = "0"
        End If
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Tipo de Subsidio Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox4.Text) = 0 Then
            Me.TextBox4.Text = "0"
        End If
        If Len(Me.TextBox5.Text) = 0 Then
            Me.TextBox5.Text = "0"
        End If
        If Len(Me.TextBox6.Text) = 0 Then
            Me.TextBox6.Text = "."
        End If
        If Len(Me.TextBox7.Text) = 0 Then
            Me.TextBox7.Text = "0"
        End If
        If Len(Me.ComboBox2.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar la Especialidad Medica Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox8.Text) = 0 Then
            Me.TextBox8.Text = "."
        End If
        If Len(Me.TextBox9.Text) = 0 Then
            Me.TextBox9.Text = "."
        End If
        Call EDITAR_REGISTRO()
        vmssID_SUBS = Me.TextBox1.Text
        CERRAR = False
        Me.Close()
    End Sub
    Private Sub EDITAR_REGISTRO()
        Dim fFECHA_I As String
        If Me.DateTimePicker1.Checked = True Then
            fFECHA_I = Mid(Me.DateTimePicker1.Value, 1, 10)
        Else
            fFECHA_I = "NULL"
        End If
        Dim f1 As Object = If(fFECHA_I <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_I + "', 105), 23)", "NULL")

        Dim fFECHA_F As String
        If Me.DateTimePicker2.Checked = True Then
            fFECHA_f = Mid(Me.DateTimePicker2.Value, 1, 10)
        Else
            fFECHA_f = "NULL"
        End If
        Dim f2 As Object = If(fFECHA_F <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_F + "', 105), 23)", "NULL")

        Dim fFECHA_PARTO As String
        If Me.DateTimePicker3.Checked = True Then
            fFECHA_PARTO = Mid(Me.DateTimePicker3.Value, 1, 10)
        Else
            fFECHA_PARTO = "NULL"
        End If
        Dim f3 As Object = If(fFECHA_PARTO <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_PARTO + "', 105), 23)", "NULL")

        Dim fFECHA_PARTO_PROB As String
        If Me.DateTimePicker4.Checked = True Then
            fFECHA_PARTO_PROB = Mid(Me.DateTimePicker4.Value, 1, 10)
        Else
            fFECHA_PARTO_PROB = "NULL"
        End If
        Dim f4 As Object = If(fFECHA_PARTO_PROB <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_PARTO_PROB + "', 105), 23)", "NULL")

        Dim fFECHA_ACC_LAB As String
        If Me.DateTimePicker5.Checked = True Then
            fFECHA_ACC_LAB = Mid(Me.DateTimePicker5.Value, 1, 10)
        Else
            fFECHA_ACC_LAB = "NULL"
        End If
        Dim f5 As Object = If(fFECHA_ACC_LAB <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_ACC_LAB + "', 105), 23)", "NULL")

        Dim fFECHA_DECLARACION As String
        If Me.DateTimePicker6.Checked = True Then
            fFECHA_DECLARACION = Mid(Me.DateTimePicker6.Value, 1, 10)
        Else
            fFECHA_DECLARACION = "NULL"
        End If
        Dim f6 As Object = If(fFECHA_DECLARACION <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_DECLARACION + "', 105), 23)", "NULL")
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE SUBSIDIOS] SET NO_EXPEDIENTE = '" & Me.TextBox2.Text & "', NO_BOLETA = '" & Me.TextBox3.Text & "', ID_T_SUBS = " & Me.ComboBox1.SelectedValue & ", N_ORDEN = '" & Me.TextBox4.Text & "', FECHA_I = " & f1 & ", FECHA_F = " & f2 & ", CANT_DIAS = " & Me.TextBox5.Text & ", FECHA_PARTO = " & f3 & " , FECHA_PARTO_PROB = " & f4 & " , FECHA_ACC_LAB = " & f5 & " , FECHA_DECLARACION = " & f6 & ", VALOR_X_DIA = 0, VALOR_TOTAL = 0, MEDICO = '" & Me.TextBox6.Text & "', N_MINSA = '" & Me.TextBox7.Text & "', ID_E_MED = " & Me.ComboBox2.SelectedValue & ", DIAGNOSTICO = '" & Me.TextBox8.Text & "', OBSERVACIONES = '" & Me.TextBox9.Text & "' WHERE ID_SUBS = " & Me.TextBox1.Text & ""
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