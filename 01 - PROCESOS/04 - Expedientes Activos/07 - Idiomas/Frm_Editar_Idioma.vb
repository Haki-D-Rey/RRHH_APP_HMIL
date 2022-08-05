Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Editar_Idioma
    Private Sub Frm_Editar_Idioma_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox3.Text = ""
            Me.CheckBox2.Checked = False
            Me.TextBox2.Text = ""
            Me.CheckBox3.Checked = False
            Me.TextBox4.Text = ""
            Me.CheckBox3.Checked = False
            Me.TextBox5.Text = ""
            Me.TextBox1.Text = ""
            Me.ComboBox1.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Editar_Idioma_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = vmiID_M_I
        Call CARGAR_COMBO_CAT_DE_IDIOMAS()
        Me.ComboBox1.Text = vmiD_IDIOMA
        Me.CheckBox1.Checked = vmiHABLA
        Me.TextBox2.Text = vmiP_HABLA
        Me.CheckBox2.Checked = vmiLEE
        Me.TextBox3.Text = vmiPORCENTAJE_L
        Me.CheckBox3.Checked = vmiESCRIBE
        Me.TextBox4.Text = vmiPORCENTAJE_E
        Me.TextBox5.Text = vmiOBSERVACIONES
        Me.ComboBox1.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox3.Text = ""
        Me.CheckBox2.Checked = False
        Me.TextBox2.Text = ""
        Me.CheckBox3.Checked = False
        Me.TextBox4.Text = ""
        Me.CheckBox3.Checked = False
        Me.TextBox5.Text = ""
        Me.TextBox1.Text = ""
        Me.ComboBox1.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_IDIOMAS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE IDIOMAS] order by ID_IDIOMA", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_IDIOMA"
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
    Private Sub CheckBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub CheckBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub CheckBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox3.KeyDown
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
            MsgBox("Debe seleccionar el Idioma Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If (Me.CheckBox1.Checked = True And Len(Me.TextBox2.Text) = 0) Or (Me.CheckBox1.Checked = True And Me.TextBox2.Text = "0") Then
            MsgBox("Si no hay Porcentaje existente debe desmarcar la casilla <HABLA>", vbInformation, "Mensaje del Sistema")
            Me.TextBox2.Focus()
            Exit Sub
        End If
        If (Me.CheckBox2.Checked = True And Len(Me.TextBox3.Text) = 0) Or (Me.CheckBox2.Checked = True And Me.TextBox3.Text = "0") Then
            MsgBox("Si no hay Porcentaje existente debe desmarcar la casilla <LEE>", vbInformation, "Mensaje del Sistema")
            Me.TextBox3.Focus()
            Exit Sub
        End If
        If (Me.CheckBox3.Checked = True And Len(Me.TextBox4.Text) = 0) Or (Me.CheckBox3.Checked = True And Me.TextBox4.Text = "0") Then
            MsgBox("Si no hay Porcentaje existente debe desmarcar la casilla <ESCRIBE>", vbInformation, "Mensaje del Sistema")
            Me.TextBox4.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox5.Text) = 0 Then
            Me.TextBox5.Text = "."
        End If
        Call EDITAR_REGISTRO()
        vmiID_M_I = Me.TextBox1.Text
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
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE IDIOMAS] SET ID_IDIOMA = " & Me.ComboBox1.SelectedValue & ", HABLA = '" & Me.CheckBox1.Checked & "', P_HABLA = '" & Me.TextBox2.Text & "', LEE = '" & Me.CheckBox2.Checked & "', PORCENTAJE_L = '" & Me.TextBox3.Text & "', ESCRIBE = '" & Me.CheckBox3.Checked & "', PORCENTAJE_E = '" & Me.TextBox4.Text & "', OBSERVACIONES = '" & Me.TextBox5.Text & "' WHERE ID_M_I = " & Me.TextBox1.Text & ""
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