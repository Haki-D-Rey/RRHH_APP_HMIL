Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Accesos_Asignar
    Private Sub Frm_Accesos_Asignar_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.ComboBox1.Focus()
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Accesos_Asignar_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = uUSUARIO
        Me.TextBox1.ReadOnly = True
        Call CARGAR_COMBO_MAESTRO_DE_USUARIOS()
        Me.ComboBox1.Focus()
    End Sub
    Private Sub CARGAR_COMBO_MAESTRO_DE_USUARIOS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [VISTA MAESTRO DE USUARIOS] WHERE (ID_TIPO_USUARIO = " & Frm_Usuarios.ComboBox2.SelectedValue & ") ORDER BY USUARIO", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "USUARIO"
                .ValueMember = "ID_USUARIO"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.ComboBox1.Focus()
        Me.Close()
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.TextBox1.Text = "" Then
            MsgBox("No se encuentra el Usuario de Origen", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Usuario de Destino Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Esta seguro de Asignar los Permisos del Usuario " & Me.TextBox1.Text & " al Usuario " & Me.ComboBox1.Text & "?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            Call ELIMINAR_MAESTRO_DE_NODOS_USUARIOS()
            Me.Cursor = Cursors.WaitCursor
            Call TRASLADAR_TV()
            MsgBox("Asignacion Satisfactoria", vbInformation, "Mensaje del Sistema")
            Me.Close()
            Me.Cursor = Cursors.Default
        End If
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
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [MAESTRO DE NODOS USUARIOS] WHERE ID_USUARIO = " & uID_USUARIO & " ORDER BY ID_NODO", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                'Call ELIMINAR_TODO()
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
                                       "values (" & tID_NODO & ", '" & tDESCRIPCION & "', '" & tID_NODO_PADRE & "', " & tACCESO & ", '" & tCLAVE & "', " & Me.ComboBox1.SelectedValue & ")"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ELIMINAR_MAESTRO_DE_NODOS_USUARIOS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE NODOS USUARIOS] WHERE ID_USUARIO = " & CInt(Me.ComboBox1.SelectedValue) & ""
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