Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Editar_Nucleo_Familiar
    Private Sub Frm_Editar_Nucleo_Familiar_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = ""
            Me.CheckBox5.Checked = False
            Me.TextBox2.Text = ""
            Me.DateTimePicker1.Value = Now
            Me.TextBox3.Text = ""
            Me.TextBox4.Text = ""
            Me.TextBox5.Text = ""
            Me.TextBox6.Text = ""
            Me.TextBox7.Text = ""
            Me.TextBox8.Text = ""
            Me.ComboBox1.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Editar_Nucleo_Familiar_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = vmnfID_NF
        Call CARGAR_COMBO_CAT_DE_PARENTELA()
        Me.ComboBox1.Text = vmnfD_PARENTELA
        Me.CheckBox5.Checked = vmnfFALLECIDO
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_SEXO()
        Me.ComboBox2.Text = vmnfD_T_SEXO
        Me.TextBox2.Text = vmnfAPELLIDOS_NOMBRES

        If vmnfFEC_NAC <> "" Then
            Me.DateTimePicker1.Value = vmnfFEC_NAC
            Me.DateTimePicker1.Checked = True
        Else
            Me.DateTimePicker1.Checked = False
        End If

        Call CARGAR_COMBO_CAT_DE_NIVEL_ACADEMICO()
        Me.ComboBox3.Text = vmnfD_N_ACADEMICO
        Me.TextBox3.Text = vmnfDIRECCION_D
        Me.TextBox4.Text = vmnfTELEFONO_1
        Me.TextBox5.Text = vmnfTELEFONO_2
        Me.TextBox6.Text = vmnfLUGAR_T
        Me.TextBox7.Text = vmnfDIRECCION_T
        Me.TextBox8.Text = vmnfCARGO_T
        Me.ComboBox1.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox1.Text = ""
        Me.CheckBox5.Checked = False
        Me.TextBox2.Text = ""
        Me.DateTimePicker1.Value = Now
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Me.TextBox6.Text = ""
        Me.TextBox7.Text = ""
        Me.TextBox8.Text = ""
        Me.ComboBox1.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_PARENTELA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE PARENTELA] order by ID_PARENTELA", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_PARENTELA"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_SEXO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE SEXO] order by ID_T_SEXO", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox2
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_T_SEXO"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_NIVEL_ACADEMICO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE NIVEL ACADEMICO] order by ID_N_ACADEMICO", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox3
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_N_ACADEMICO"
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
    Private Sub CheckBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox5.KeyDown
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
    Private Sub DateTimePicker1_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox3.KeyDown
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "0" Or TextBox1.Text = "" Then
            MsgBox("No se ha generado el Codigo del registro", vbInformation, "Mensaje del Sistema")
            Me.Button2.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Tipo de Parentela Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox2.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Tipo de Sexo Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox2.Text) = 0 Then
            MsgBox("Debe digitar los Apellidos y Nombres de la Parentela", vbInformation, "Mensaje del Sistema")
            Me.TextBox2.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox3.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Nivel Academico Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox3.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox3.Text) = 0 Then
            MsgBox("Debe digitar la Dirección Domiciliar de la Parentela", vbInformation, "Mensaje del Sistema")
            Me.TextBox3.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox4.Text) = 0 Then
            Me.TextBox4.Text = ""
        End If
        If Len(Me.TextBox5.Text) = 0 Then
            Me.TextBox5.Text = ""
        End If
        If Len(Me.TextBox6.Text) = 0 Then
            Me.TextBox6.Text = ""
        End If
        If Len(Me.TextBox7.Text) = 0 Then
            Me.TextBox7.Text = ""
        End If
        If Len(Me.TextBox8.Text) = 0 Then
            Me.TextBox8.Text = ""
        End If
        Call EDITAR_REGISTRO()
        vmnfID_NF = Me.TextBox1.Text
        CERRAR = False
        Me.Close()
    End Sub
    Private Sub EDITAR_REGISTRO()
        Dim fFEC_NAC As String
        If Me.DateTimePicker1.Checked = True Then
            fFEC_NAC = Mid(Me.DateTimePicker1.Value, 1, 10)
        Else
            fFEC_NAC = "NULL"
        End If
        Dim fechaNacimiento As Object = If(fFEC_NAC <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFEC_NAC + "', 105), 23)", "NULL")
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE NUCLEO FAMILIAR] SET ID_PARENTELA = " & Me.ComboBox1.SelectedValue & ", APELLIDOS_NOMBRES = '" & Trim(Me.TextBox2.Text) & "', ID_T_SEXO = " & Me.ComboBox2.SelectedValue & ", FEC_NAC = " & fechaNacimiento & ", ID_N_ACADEMICO = " & Me.ComboBox3.SelectedValue & ", DIRECCION_D = '" & Me.TextBox3.Text & "', TELEFONO_1 = '" & Me.TextBox4.Text & "', TELEFONO_2 = '" & Me.TextBox5.Text & "', LUGAR_T = '" & Me.TextBox6.Text & "', DIRECCION_T = '" & Me.TextBox7.Text & "', CARGO_T = '" & Me.TextBox8.Text & "', FALLECIDO = '" & Me.CheckBox5.Checked & "' WHERE ID_NF = " & Me.TextBox1.Text & ""
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