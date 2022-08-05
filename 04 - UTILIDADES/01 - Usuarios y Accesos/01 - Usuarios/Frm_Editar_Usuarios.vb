Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Editar_Usuarios
    Private Sub Frm_Editar_Usuarios_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Editar_Usuarios_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call CARGAR_COMBO_SESION()
        Call CARGAR_COMBO_ESTADO_USUARIO()
        Call CARGAR_COMBO_CAT_DE_TIPOS_DE_USUARIO()
        Me.TextBox4.Text = uID_USUARIO
        Me.TextBox1.Text = uUSUARIO
        Me.TextBox2.Text = uCUENTA
        Me.TextBox3.Text = uCLAVE
        Me.ComboBox1.Text = uD_SESION_USUARIO
        Me.ComboBox2.Text = uD_ESTADO_USUARIO
        Me.ComboBox3.Text = uD_TIPO_USUARIO
        Me.TextBox5.Text = uCODIGO
        Me.TextBox6.Text = uCARGO
        Me.TextBox7.Text = uNIVEL1
        Me.TextBox8.Text = uNIVEL2
        Me.TextBox9.Text = uNIVEL3
        Me.TextBox10.Text = uNIVEL4
        Me.TextBox11.Text = uNIVEL5
        Me.TextBox12.Text = uNIVEL6
        Me.TextBox13.Text = uNIVEL7
        Me.TextBox14.Text = uNIVEL8
        Me.TextBox1.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.TextBox1.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_TIPOS_DE_USUARIO()
        If isuID_TIPO_USUARIO = 1 Then
            CADENAsqlCombo = "Select * FROM [CAT DE TIPO DE USUARIO] ORDER BY DESCRIPCION"
        Else
            CADENAsqlCombo = "SELECT * FROM [CAT DE TIPO DE USUARIO] WHERE ID_TIPO_USUARIO <> 1 ORDER BY DESCRIPCION"
        End If
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter(CADENAsqlCombo, Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox3
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_TIPO_USUARIO"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_ESTADO_USUARIO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTADO DE USUARIO] order by ID_ESTADO_USUARIO", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox2
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_ESTADO_USUARIO"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_SESION()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE SESION DE USUARIO] order by ID_SESION_USUARIO", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_SESION_USUARIO"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox1.CheckedChanged
        Call MOSTRAR_CONTRASEÑA()
    End Sub
    Private Sub MOSTRAR_CONTRASEÑA()
        If Me.CheckBox1.Checked = True Then
            Me.TextBox3.PasswordChar = ""
        Else
            Me.TextBox3.PasswordChar = "•"
        End If
    End Sub
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
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
    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
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
    Private Sub ComboBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox3.KeyDown
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
    Private Sub TextBox12_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox12.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox13_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox13.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox14_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox14.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox1_DoubleClick(sender As Object, e As EventArgs) Handles TextBox1.DoubleClick
        Frm_Buscar_en_RRHH.ShowDialog()
        If DATO_EN_RRHH = True Then
            Me.TextBox1.Text = bEMPLEADO
            Me.TextBox5.Text = bCODIGO
            Me.TextBox6.Text = bCARGO
            Me.TextBox7.Text = bD_N1
            Me.TextBox8.Text = bD_N2
            Me.TextBox9.Text = bD_N3
            Me.TextBox10.Text = bD_N4
            Me.TextBox11.Text = bD_N5
            Me.TextBox12.Text = bD_N6
            Me.TextBox13.Text = bD_N7
            Me.TextBox14.Text = bD_N8
            Me.TextBox2.Text = bCUENTA
            Me.TextBox3.Focus()
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox4.Text = "0" Or TextBox4.Text = "" Then
            MsgBox("No se ha generado el Codigo del registro", vbInformation, "Mensaje del Sistema")
            Me.TextBox4.Text = "0"
            Me.Button2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox1.Text) = 0 Then
            MsgBox("Debe digitar los Nombres y Apellidos del Usuario", vbInformation, "Mensaje del Sistema")
            Me.TextBox1.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox2.Text) = 0 Then
            MsgBox("Debe digitar la Cuenta del Usuario", vbInformation, "Mensaje del Sistema")
            Me.TextBox2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox3.Text) = 0 Then
            MsgBox("Debe digitar la Contraseña del Usuario", vbInformation, "Mensaje del Sistema")
            Me.TextBox3.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Estado Inicial de la Sesion del Usuario Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox2.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Estado del Usuario Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox5.Text) = 0 Then
            Me.TextBox5.Text = "."
        End If
        If Len(Me.TextBox6.Text) = 0 Then
            Me.TextBox6.Text = "."
        End If
        If Len(Me.TextBox7.Text) = 0 Then
            Me.TextBox7.Text = "."
        End If
        Call EDITAR_REGISTRO()
        uID_USUARIO = Me.TextBox4.Text
        Me.TextBox4.Text = "0"
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.ComboBox1.Text = "NO ACTIVA"
        Me.ComboBox2.Text = "ACTIVO"
        Me.TextBox5.Text = ""
        Me.TextBox6.Text = ""
        Me.TextBox7.Text = ""

        Me.TextBox1.Focus()
        CERRAR = False
        Me.Close()
    End Sub
    Private Sub EDITAR_REGISTRO()
        Dim fFECHA_I1 As String
        Call Clases.BUSCAR_FECHA_Y_HORA()
        fFECHA_I1 = Mid(FECHA_SERVIDOR, 1, 10)
        Dim f1 As Object = If(fFECHA_I1 <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_I1 + "', 105), 23)", "NULL")
        Dim vIP As String = Clases.GetIP()
        If Me.ComboBox2.Text = "NO ACTIVO" Then
            CADENAsql = "UPDATE [MAESTRO DE USUARIOS] SET ID_TIPO_USUARIO = " & Me.ComboBox3.SelectedValue & ", USUARIO = '" & Me.TextBox1.Text & "', CUENTA  = '" & Me.TextBox2.Text & "', ID_SESION_USUARIO = " & Me.ComboBox1.SelectedValue & ", ID_ESTADO_USUARIO = " & Me.ComboBox2.SelectedValue & ", CODIGO = '" & Me.TextBox5.Text & "', CARGO = '" & Me.TextBox6.Text & "', NIVEL1 = '" & Me.TextBox7.Text & "', NIVEL2 = '" & Me.TextBox8.Text & "', NIVEL3 = '" & Me.TextBox9.Text & "', NIVEL4 = '" & Me.TextBox10.Text & "', NIVEL5 = '" & Me.TextBox11.Text & "', NIVEL6 = '" & Me.TextBox12.Text & "', NIVEL7 = '" & Me.TextBox13.Text & "', NIVEL8 = '" & Me.TextBox14.Text & "', FECHA_BAJA = " & f1 & ", USUARIO_EJECUTA_BAJA = '" & isuCUENTA & "', TERMINAL_BAJA = '" & My.Computer.Name & "', IP_BAJA = '" & vIP & "' WHERE ID_USUARIO = " & CInt(Me.TextBox4.Text) & ""
        Else
            CADENAsql = "UPDATE [MAESTRO DE USUARIOS] SET ID_TIPO_USUARIO = " & Me.ComboBox3.SelectedValue & ", USUARIO = '" & Me.TextBox1.Text & "', CUENTA  = '" & Me.TextBox2.Text & "', ID_SESION_USUARIO = " & Me.ComboBox1.SelectedValue & ", ID_ESTADO_USUARIO = " & Me.ComboBox2.SelectedValue & ", CODIGO = '" & Me.TextBox5.Text & "', CARGO = '" & Me.TextBox6.Text & "', NIVEL1 = '" & Me.TextBox7.Text & "', NIVEL2 = '" & Me.TextBox8.Text & "', NIVEL3 = '" & Me.TextBox9.Text & "', NIVEL4 = '" & Me.TextBox10.Text & "', NIVEL5 = '" & Me.TextBox11.Text & "', NIVEL6 = '" & Me.TextBox12.Text & "', NIVEL7 = '" & Me.TextBox13.Text & "', NIVEL8 = '" & Me.TextBox14.Text & "' WHERE ID_USUARIO = " & CInt(Me.TextBox4.Text) & ""
        End If
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = CADENAsql
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