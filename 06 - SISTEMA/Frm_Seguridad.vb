Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Seguridad
    Dim EXISTE As Boolean = False
    Private Sub Frm_Seguridad_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox2.Text = ""
            Me.TextBox1.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Seguridad_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = Environment.UserName
        Me.TextBox2.Text = ""
        Me.TextBox1.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.TextBox2.Text = ""
        Me.TextBox1.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub TextBox1_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            Me.TextBox2.Focus()
        End If
    End Sub
    Private Sub TextBox2_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            Call VALIDAR()
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call VALIDAR()
    End Sub
    Private Sub VALIDAR()
        If Me.TextBox1.Text = "" Then
            MsgBox("Debe digitar la Cuenta del Usuario Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox1.Focus()
            Exit Sub
        End If
        If Me.TextBox2.Text = "" Then
            MsgBox("Debe digitar la Contraseña del Usuario Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox2.Focus()
            Exit Sub
        End If
        Call BUSCAR_CUENTA_USUARIO()
        If EXISTE = False Then
            Exit Sub
        Else
            CERRAR = False
            Me.Close()
        End If
    End Sub
    Dim ES_ADMINISTRADOR As Boolean
    Private Sub BUSCAR_CUENTA_USUARIO()
        Dim xSqlDataAdapter As SqlDataAdapter
        Dim xDataTable As New DataTable
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Try
            Conexion.ABD_RRHH()
            xSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [VISTA MAESTRO DE USUARIOS] WHERE CUENTA = '" & UCase(Trim(Me.TextBox1.Text)) & "' AND CLAVE = '" & ENCRIPTAR.ENCRIPTAR_PALABRA(Me.TextBox2.Text) & "' ORDER BY ID_USUARIO", Conexion.CxRRHH)
            xDataTable.Clear()
            xSqlDataAdapter.Fill(xDataTable)
            If xDataTable.Rows.Count > 0 Then
                isuID_USUARIO = xDataTable.Rows(0).Item(0).ToString
                isuID_TIPO_USUARIO = xDataTable.Rows(0).Item(1).ToString
                isuD_TIPO_USUARIO = xDataTable.Rows(0).Item(2).ToString
                isuUSUARIO = xDataTable.Rows(0).Item(3).ToString
                If isuID_TIPO_USUARIO = 5 Then
                    Call BUSCAR_ADMINISTRADOR_DE_RECURSOS()
                End If
                isuCUENTA = xDataTable.Rows(0).Item(4).ToString
                isuCLAVE = xDataTable.Rows(0).Item(5).ToString
                isuID_SESION_USUARIO = xDataTable.Rows(0).Item(6).ToString
                isuD_SESION_USUARIO = xDataTable.Rows(0).Item(7).ToString
                isuID_ESTADO_USUARIO = xDataTable.Rows(0).Item(8).ToString
                isuD_ESTADO_USUARIO = xDataTable.Rows(0).Item(9).ToString
                isuCODIGO = xDataTable.Rows(0).Item(10).ToString
                isuCARGO = xDataTable.Rows(0).Item(11).ToString
                isuN1 = xDataTable.Rows(0).Item(12).ToString
                isuN2 = xDataTable.Rows(0).Item(13).ToString
                isuN3 = xDataTable.Rows(0).Item(14).ToString
                isuN4 = xDataTable.Rows(0).Item(15).ToString
                isuN5 = xDataTable.Rows(0).Item(16).ToString
                isuN6 = xDataTable.Rows(0).Item(17).ToString
                isuN7 = xDataTable.Rows(0).Item(18).ToString
                isuN8 = xDataTable.Rows(0).Item(19).ToString
                If isuID_USUARIO = 2 Or isuID_USUARIO = 3 Then  'YARID.ARIAS : 2, GILMARDES.CALVO : 3
                    Call ACTUALIZAR_ACCESO_DE_USUARIO_CON_PERFIL_ADMIN()
                End If
                If isuID_SESION_USUARIO = 1 Then
                    EXISTE = False
                    MsgBox("Ya existe un Inicio de Sesion para esta Cuenta o no se ha actualizado la politica de seguridad, intentelo nuevamente", vbInformation, "Mensaje del Sistema")
                    Me.TextBox1.Focus()
                    Me.TextBox1.SelectAll()
                    Exit Sub
                End If
                If isuID_ESTADO_USUARIO = 2 Then
                    EXISTE = False
                    MsgBox("La Cuenta se encuentra Deshabilitada", vbInformation, "Mensaje del Sistema")
                    Me.TextBox1.Focus()
                    Me.TextBox1.SelectAll()
                    Exit Sub
                End If
                EXISTE = True
                If isuID_TIPO_USUARIO > 2 Then
                    Call Me.ACTUALIZAR_REGISTRO()
                End If
            Else
                EXISTE = False
                MsgBox("No se encuentra el Usuario o la Cuenta y Contraseña son incorrectas", vbInformation, "Mensaje del Sistema")
                Me.TextBox1.Focus()
                Me.TextBox1.SelectAll()
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub BUSCAR_ADMINISTRADOR_DE_RECURSOS()
        Dim xSqlDataAdapterU As SqlDataAdapter
        Dim xDataTableU As New DataTable
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Try
            Conexion.ABD_RRHH()
            xSqlDataAdapterU = New SqlDataAdapter("SELECT * FROM [CAT DE ADMINISTRADORES] WHERE DESCRIPCION = '" & UCase(Trim(isuUSUARIO)) & "' ORDER BY DESCRIPCION", Conexion.CxRRHH)
            xDataTableU.Clear()
            xSqlDataAdapterU.Fill(xDataTableU)
            If xDataTableU.Rows.Count > 0 Then
                isuUSUARIO_R = xDataTableU.Rows(0).Item(0).ToString
            Else
                isuUSUARIO_R = 0
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ACTUALIZAR_ACCESO_DE_USUARIO_CON_PERFIL_ADMIN()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE USUARIOS] SET ID_SESION_USUARIO = '2' WHERE ID_USUARIO = " & CInt(isuID_USUARIO) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ACTUALIZAR_REGISTRO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE USUARIOS] SET ID_SESION_USUARIO = '1' WHERE ID_USUARIO = " & CInt(isuID_USUARIO) & ""
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