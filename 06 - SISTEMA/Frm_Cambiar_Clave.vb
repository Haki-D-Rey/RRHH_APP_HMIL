Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Cambiar_Clave
    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        CERRAR = True
        Me.TextBox3.Focus()
        On Error GoTo PROXIMA_LINEA
PROXIMA_LINEA:
        Me.Close()
    End Sub
    Private Sub Frm_Cambiar_Clave_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            CERRAR = True
            Me.TextBox3.Focus()
            On Error GoTo PROXIMA_LINEA
PROXIMA_LINEA:
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Cambiar_Clave_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Call BUSCAR_USUARIO()
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox3.Focus()
    End Sub
    Dim CONTRA_ACTUAL As String
    Private Sub BUSCAR_USUARIO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [MAESTRO DE USUARIOS] WHERE ID_USUARIO = " & isuID_USUARIO & " ORDER BY ID_USUARIO", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                Me.TextBox6.Text = MiDataTable.Rows(0).Item(0).ToString
                Me.TextBox1.Text = MiDataTable.Rows(0).Item(2).ToString
                Me.TextBox2.Text = MiDataTable.Rows(0).Item(3).ToString
                CONTRA_ACTUAL = MiDataTable.Rows(0).Item(4).ToString
            Else
                Me.TextBox6.Text = "0"
                Me.TextBox1.Text = ""
                Me.TextBox2.Text = ""
                CONTRA_ACTUAL = ""
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim PCONTRA As Integer
    Private Sub CheckBox1_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox1.CheckedChanged
        PCONTRA = 1
        Call MOSTRAR_CONTRASEÑA()
    End Sub
    Private Sub MOSTRAR_CONTRASEÑA()
        If PCONTRA = 1 Then
            If Me.CheckBox1.Checked = True Then
                Me.TextBox3.PasswordChar = ""
            Else
                Me.TextBox3.PasswordChar = "•"
            End If
        End If
        If PCONTRA = 2 Then
            If Me.CheckBox2.Checked = True Then
                Me.TextBox4.PasswordChar = ""
            Else
                Me.TextBox4.PasswordChar = "•"
            End If
        End If
        If PCONTRA = 3 Then
            If Me.CheckBox3.Checked = True Then
                Me.TextBox5.PasswordChar = ""
            Else
                Me.TextBox5.PasswordChar = "•"
            End If
        End If
    End Sub
    Private Sub CheckBox2_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox2.CheckedChanged
        PCONTRA = 2
        Call MOSTRAR_CONTRASEÑA()
    End Sub
    Private Sub CheckBox3_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox3.CheckedChanged
        PCONTRA = 3
        Call MOSTRAR_CONTRASEÑA()
    End Sub
    Private Sub TextBox3_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox4_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox4.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox5_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox5.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If Me.TextBox6.Text = "" Or Me.TextBox6.Text = "0" Then 'CODIGO DEL REGISTRO
            MsgBox("No se encuentra el codigo del registro", vbInformation, "Mensaje del Sistema")
            Me.Button2.Focus()
            Exit Sub
        End If
        If Me.TextBox1.Text = "" Then   'NOMBRE DEL USUARIO
            MsgBox("No se encuentra el Nombre del Usuario", vbInformation, "Mensaje del Sistema")
            Me.Button2.Focus()
            Exit Sub
        End If
        If Me.TextBox2.Text = "" Then   'NOMBRE DE LA CUENTA
            MsgBox("No se encuentra el Nombre de Cuenta", vbInformation, "Mensaje del Sistema")
            Me.Button2.Focus()
            Exit Sub
        End If
        If ENCRIPTAR.ENCRIPTAR_PALABRA(Me.TextBox3.Text) <> UCase(Trim(CONTRA_ACTUAL)) Then 'CONTRASEÑA ACTUAL
            MsgBox("La Contraseña Actual no es válida", vbInformation, "Mensaje del Sistema")
            Me.TextBox3.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox4.Text) < 6 Then   'NUEVA CONTRASEÑA
            MsgBox("La Nueva Contraseña no puede ser menor a 6 caracteres", vbInformation, "Mensaje del Sistema")
            Me.TextBox3.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox4.Text) > 10 Then
            MsgBox("La Nueva Contraseña no puede ser Mayor a 10 caracteres", vbInformation, "Mensaje del Sistema")
            Me.TextBox4.Focus()
            Exit Sub
        End If
        If UCase(Trim(Me.TextBox4.Text)) = UCase(Trim(Me.TextBox3.Text)) Then '
            MsgBox("La Contraseña Nueva no puede ser igual a la Contraseña Actual", vbInformation, "Mensaje del Sistema")
            Me.TextBox4.Focus()
            Exit Sub
        End If
        If UCase(Trim(Me.TextBox4.Text)) <> UCase(Trim(Me.TextBox5.Text)) Then 'CONFIRMAR CONTRASEÑA
            MsgBox("La Confirmacion de la Contraseña es incorrecta", vbInformation, "Mensaje del Sistema")
            Me.TextBox5.Focus()
            Exit Sub
        End If
        Call EDITAR_REGISTRO()
        Me.TextBox3.Focus()
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
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE USUARIOS] SET CLAVE = '" & ENCRIPTAR.ENCRIPTAR_PALABRA(Me.TextBox4.Text) & "' WHERE ID_USUARIO = " & CInt(Me.TextBox6.Text) & ""
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