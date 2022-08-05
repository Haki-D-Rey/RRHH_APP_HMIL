Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Agregar_Usuarios
    Private Sub Frm_Agregar_Usuarios_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Agregar_Usuarios_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call CARGAR_COMBO_SESION()
        Call CARGAR_COMBO_ESTADO_USUARIO()
        Call CARGAR_COMBO_CAT_DE_TIPOS_DE_USUARIO()
        Me.TextBox4.Text = "0"
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.ComboBox1.Text = "NO ACTIVA"
        Me.ComboBox2.Text = "ACTIVO"
        Me.ComboBox3.Text = Frm_Usuarios.ComboBox2.Text
        Me.TextBox5.Text = ""
        Me.TextBox6.Text = ""
        Me.TextBox7.Text = ""
        Me.TextBox8.Text = ""
        Me.TextBox9.Text = ""
        Me.TextBox10.Text = ""
        Me.TextBox11.Text = ""
        Me.TextBox12.Text = ""
        Me.TextBox13.Text = ""
        Me.TextBox14.Text = ""
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
    Private Sub GENERAR_CODIGO()
        CODIGO = 0
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            MiDataAdapter = New SqlDataAdapter("SELECT ID_USUARIO FROM [MAESTRO DE USUARIOS] ORDER BY ID_USUARIO", Conexion.CxRRHH)
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
            Me.TextBox4.Text = CODIGO
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
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
    Private Sub GRABAR_ACCESO()
        Me.Cursor = Cursors.WaitCursor
        Call TRASLADAR_TV()
        Me.Cursor = Cursors.Default
    End Sub
    Dim tID_NODO As Integer
    Dim tDESCRIPCION As String
    Dim tID_NODO_PADRE As String
    Dim tACCESO As Integer
    Dim tCLAVE As String
    Dim CODIGO As Integer
    Private Sub TRASLADAR_TV()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [MAESTRO DE NODOS BASE] ORDER BY ID_NODO", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                Call ELIMINAR_TODO()
                For Ix = 0 To MiDataTable.Rows.Count - 1
                    tID_NODO = MiDataTable.Rows(Ix).Item(0).ToString
                    tDESCRIPCION = MiDataTable.Rows(Ix).Item(1).ToString
                    tID_NODO_PADRE = MiDataTable.Rows(Ix).Item(2).ToString
                    tACCESO = MiDataTable.Rows(Ix).Item(3).ToString
                    tCLAVE = MiDataTable.Rows(Ix).Item(4).ToString
                    Call GRABAR()
                Next
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub GRABAR()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE NODOS USUARIOS] (ID_NODO, DESCRIPCION, ID_NODO_PADRE, ACCESO, CLAVE, ID_USUARIO) " &
                                       "values (" & tID_NODO & ", '" & tDESCRIPCION & "', '" & tID_NODO_PADRE & "', " & tACCESO & ", '" & tCLAVE & "', " & CInt(CODIGO) & ")"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ELIMINAR_TODO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE NODOS USUARIOS] WHERE ID_USUARIO = " & CInt(CODIGO) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.TextBox4.Text = "0"
        Call GENERAR_CODIGO()
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
        Call GRABAR_REGISTRO()
        Call GRABAR_ACCESO()
        '-------------------------------------------
        If Me.ComboBox3.Text = "OPERADOR EXTERNO" Then
            Call BUSCAR_SI_EXISTE_EN_CAT_DE_ADMINISTRADORES()
            If EXISTE_ADMIN = True Then
                MsgBox("Ya existe un registro con esta Descripcion, el registro actual no sera ingresado al Catalogo de <Administradores de Recursos>", vbInformation, "Mensaje del Sistema")
            Else
                Call GENERAR_ID_ADMIN()
                Call AGREGAR_REGISTRO_ADMIN
            End If
        End If
        '-------------------------------------------
        uID_USUARIO = Me.TextBox4.Text
        Me.TextBox4.Text = "0"
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.ComboBox1.Text = "NO ACTIVA"
        Me.ComboBox2.Text = "ACTIVO"
        Me.ComboBox3.Text = Frm_Usuarios.ComboBox2.Text
        Me.TextBox5.Text = ""
        Me.TextBox6.Text = ""
        Me.TextBox7.Text = ""
        Me.TextBox1.Focus()
        CERRAR = False
        Me.Close()
    End Sub
    Private Sub AGREGAR_REGISTRO_ADMIN()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [CAT DE ADMINISTRADORES] (ID_ADMIN, DESCRIPCION, ACTIVO)" &
                                  "values (" & CInt(IDADMIN) & ", '" & Trim(Me.TextBox1.Text) & "', 'TRUE')"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim IDADMIN As Integer
    Private Sub GENERAR_ID_ADMIN()
        CODIGO = 0
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            MiDataAdapter = New SqlDataAdapter("SELECT ID_ADMIN FROM [CAT DE ADMINISTRADORES] ORDER BY ID_ADMIN", Conexion.CxRRHH)
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
            IDADMIN = CODIGO
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim EXISTE_ADMIN As Boolean
    Private Sub BUSCAR_SI_EXISTE_EN_CAT_DE_ADMINISTRADORES()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ADMINISTRADORES] WHERE DESCRIPCION = '" & Trim(Me.TextBox1.Text) & "' ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                EXISTE_ADMIN = True
            Else
                EXISTE_ADMIN = False
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub GRABAR_REGISTRO()
        Dim fFECHA_I1 As String
        Call Clases.BUSCAR_FECHA_Y_HORA()
        fFECHA_I1 = Mid(FECHA_SERVIDOR, 1, 10)
        Dim f1 As Object = If(fFECHA_I1 <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_I1 + "', 105), 23)", "NULL")
        Dim vIP As String = Clases.GetIP()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE USUARIOS] (ID_USUARIO, ID_TIPO_USUARIO, USUARIO, CUENTA, CLAVE, ID_SESION_USUARIO, ID_ESTADO_USUARIO, CODIGO, CARGO, NIVEL1, NIVEL2, NIVEL3, NIVEL4, NIVEL5, NIVEL6, NIVEL7, NIVEL8, FECHA_ALTA, USUARIO_EJECUTA_ALTA, TERMINAL_ALTA, IP_ALTA, FECHA_BAJA, USUARIO_EJECUTA_BAJA, TERMINAL_BAJA, IP_BAJA, ID_ESTADO_R)" &
                                  "values (" & CInt(Me.TextBox4.Text) & ", " & Me.ComboBox3.SelectedValue & ", '" & Trim(Me.TextBox1.Text) & "', '" & Trim(Me.TextBox2.Text) & "', '" & ENCRIPTAR.ENCRIPTAR_PALABRA(Me.TextBox3.Text) & "', " & Me.ComboBox1.SelectedValue & ", " & Me.ComboBox2.SelectedValue & ", '" & Trim(Me.TextBox5.Text) & "', '" & Trim(Me.TextBox6.Text) & "', '" & Trim(Me.TextBox7.Text) & "', '" & Trim(Me.TextBox8.Text) & "', '" & Trim(Me.TextBox9.Text) & "', '" & Trim(Me.TextBox10.Text) & "', '" & Trim(Me.TextBox11.Text) & "', '" & Trim(Me.TextBox12.Text) & "', '" & Trim(Me.TextBox13.Text) & "', '" & Trim(Me.TextBox14.Text) & "', " & f1 & ", '" & isuCUENTA & "', '" & My.Computer.Name & "', '" & vIP & "', NULL, NULL, NULL, NULL, 1)"
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