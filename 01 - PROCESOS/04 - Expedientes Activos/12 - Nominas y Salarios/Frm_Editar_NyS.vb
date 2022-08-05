Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Editar_NyS
    Private Sub Frm_Editar_NyS_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = ""
            Me.DateTimePicker1.Value = Now
            Me.TextBox2.Text = "0"
            Me.TextBox3.Text = "0"
            Me.CheckBox1.Checked = False
            Me.TextBox4.Text = ""
            Me.ComboBox1.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Editar_NyS_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = vmnsID_M_NYS
        Call CARGAR_COMBO_CAT_DE_MOVIMIENTO_EN_NOMINA()
        Me.ComboBox1.Text = vmnsD_MOVIMIENTO_N
        Me.DateTimePicker1.Value = vmnsF_MOVIMIENTO_N
        Call CARGAR_COMBO_CAT_DE_NOMINA()
        Me.ComboBox2.Text = vmnsD_NOMINA
        Me.TextBox2.Text = vmnsPROPUESTA
        Me.TextBox3.Text = vmnsSALARIO
        Me.CheckBox1.Checked = vmnsSAL_MINIMO
        Me.TextBox4.Text = vmnsOBSERVACIONES
        Me.CheckBox2.Checked = vmnsACTIVO
        Me.ComboBox1.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox1.Text = ""
        Me.DateTimePicker1.Value = Now
        Me.TextBox2.Text = "0"
        Me.TextBox3.Text = "0"
        Me.CheckBox1.Checked = False
        Me.TextBox4.Text = ""
        Me.ComboBox1.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_MOVIMIENTO_EN_NOMINA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE MOVIMIENTO EN NOMINA] order by ID_MOVIMIENTO_N", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_MOVIMIENTO_N"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_NOMINA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE NOMINA] order by ID_NOMINA", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox2
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_NOMINA"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
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
    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
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
    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub CheckBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox1.KeyDown
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "0" Or TextBox1.Text = "" Then
            MsgBox("No se ha generado el Codigo del registro", vbInformation, "Mensaje del Sistema")
            Me.TextBox1.Text = "0"
            Me.Button2.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Tipo de Movimiento Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox2.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar la Nómina Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox2.Text) = 0 Or Me.TextBox2.Text = "" Or Me.TextBox2.Text = "0" Then
            MsgBox("Debe digitar el monto de la Propuesta Salarial del Empleado", vbInformation, "Mensaje del Sistema")
            Me.TextBox2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox3.Text) = 0 Or Me.TextBox3.Text = "" Or Me.TextBox3.Text = "0" Then
            MsgBox("Debe digitar el monto del Salario del Empleado", vbInformation, "Mensaje del Sistema")
            Me.TextBox3.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox4.Text) = 0 Then
            Me.TextBox4.Text = ""
        End If
        If Me.CheckBox2.Checked = True Then
            Call DESACTIVAR_SALARIOS()
        End If
        Call EDITAR_REGISTRO()
        vmnsID_M_NYS = Me.TextBox1.Text
        CERRAR = False
        Me.Close()
    End Sub
    Private Sub DESACTIVAR_SALARIOS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE NOMINAS Y SALARIOS] SET SAL_ACTUAL = 'FALSE' WHERE ID_M_P = " & Frm_Expedientes_Activos.TextBox35.Text & " AND ID_NOMINA = " & Me.ComboBox2.SelectedValue & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub EDITAR_REGISTRO()
        Dim fFEC_MOV As String = Me.DateTimePicker1.Text
        Dim fechaMovimiento = "Convert(VARCHAR(10), Convert(Date, '" + fFEC_MOV + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE NOMINAS Y SALARIOS] SET ID_MOVIMIENTO_N = " & Me.ComboBox1.SelectedValue & ", F_MOVIMIENTO_N = " & fechaMovimiento & ", ID_NOMINA = " & Me.ComboBox2.SelectedValue & ", PROPUESTA = '" & Me.TextBox2.Text & "', SALARIO = '" & Me.TextBox3.Text & "', SAL_MINIMO = '" & Me.CheckBox1.Checked & "', OBSERVACIONES = '" & Me.TextBox4.Text & "', SAL_ACTUAL = '" & Me.CheckBox2.Checked & "' WHERE ID_M_NYS = " & Me.TextBox1.Text & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CheckBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
End Class