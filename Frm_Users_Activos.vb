Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Users_Activos
    Private Sub Frm_Users_Activos_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.ListView1.View = View.List
        Call CARGAR_DATOS()
    End Sub
    Private Sub CARGAR_DATOS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [MAESTRO DE USUARIOS] WHERE ID_SESION_USUARIO = 1 ORDER BY USUARIO", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                Me.ListView1.Items.Clear()
                For I = 0 To MiDataTable.Rows.Count - 1
                    Me.ListView1.Items.Add(MiDataTable.Rows(I).Item(2).ToString)
                Next
            Else
                Me.ListView1.Items.Clear()
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Frm_Users_Activos_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.F5) Then
            Call CARGAR_DATOS()
        End If
        If (e.KeyCode = Keys.Escape) Then
            Me.Close()
        End If
    End Sub
    Dim USUARIO_NO_ACTIVO As String
    Private Sub ListView1_DoubleClick(sender As Object, e As EventArgs) Handles ListView1.DoubleClick
        If ListView1.Items.Count > 0 Then
            Dim i As Integer
            For Each i In ListView1.SelectedIndices
                USUARIO_NO_ACTIVO = ListView1.Items(i).Text
                Call ACTIVAR_USUARIO()
                Call CARGAR_DATOS()
            Next
        End If
    End Sub
    Private Sub ACTIVAR_USUARIO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE USUARIOS] SET ID_SESION_USUARIO = '2' WHERE USUARIO = '" & USUARIO_NO_ACTIVO & "'"
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