Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Editar_Condecoraciones
    Private Sub Frm_Editar_Condecoraciones_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = ""
            Me.TextBox3.Text = ""
            Me.TextBox4.Text = ""
            Me.ComboBox1.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Editar_Condecoraciones_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = vmcID_M_CONDE
        Call CARGAR_COMBO_CAT_DE_CONDECORACIONES()
        Me.ComboBox1.Text = vmcD_CONDE
        Me.DateTimePicker1.Value = vmcCONDE_A
        Me.TextBox3.Text = vmcMOTIVOS
        Me.TextBox4.Text = vmcOBSERVACIONES
        Me.ComboBox1.Focus()
    End Sub

    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox1.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.ComboBox1.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_CONDECORACIONES()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE CONDECORACIONES] order by ID_CONDE", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_CONDE"
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
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs)
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "0" Or TextBox1.Text = "" Then
            MsgBox("No se ha generado el Codigo del registro", vbInformation, "Mensaje del Sistema")
            Me.TextBox1.Text = "0"
            Me.Button2.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar la Condecoración Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox3.Text) = 0 Then
            Me.TextBox3.Text = ""
        End If
        If Len(Me.TextBox4.Text) = 0 Then
            Me.TextBox4.Text = ""
        End If
        Call EDITAR_REGISTRO()
        vmcID_C_CONDE = Me.TextBox1.Text
        CERRAR = False
        Me.Close()
    End Sub
    Dim fechaCONDE As String
    Dim fFEC_CONDE As String
    Private Sub EDITAR_REGISTRO()
        fFEC_CONDE = Mid(Me.DateTimePicker1.Value, 1, 10)
        fechaCONDE = "Convert(VARCHAR(10), Convert(Date, '" + fFEC_CONDE + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE CONDECORACIONES] SET ID_CONDE = " & Me.ComboBox1.SelectedValue & ", CONDE_A = " & fechaCONDE & ", MOTIVOS = '" & Me.TextBox3.Text & "', OBSERVACIONES = '" & Me.TextBox4.Text & "' WHERE ID_M_CONDE = " & Me.TextBox1.Text & ""
            MiSqlCommand.ExecuteNonQuery()
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
End Class