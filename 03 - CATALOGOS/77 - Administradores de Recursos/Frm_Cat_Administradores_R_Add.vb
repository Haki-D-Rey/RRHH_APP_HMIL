Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Cat_Administradores_R_Add
    Private Sub Button47_Click(sender As Object, e As EventArgs) Handles Button47.Click
        Frm_Cat_Administradores_R_Add_Buscar.ShowDialog()
        If CERRAR = False Then
            Me.TextBox2.Text = bEMPLEADO
            Me.Button1.Focus()
        End If
    End Sub
    Private Sub Frm_Cat_Administradores_R_Add_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.Button47.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Cat_Administradores_R_Add_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = "0"
        Me.TextBox2.Text = ""
        Me.CheckBox1.Checked = True
        Me.CheckBox1.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Button47.Focus()
        CERRAR = True
        Me.Close()
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
            Me.TextBox1.Text = CODIGO
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.TextBox1.Text = "0"
        Call GENERAR_CODIGO()
        If TextBox1.Text = "0" Or TextBox1.Text = "" Then
            MsgBox("No se ha generado el Codigo del registro", vbInformation, "Mensaje del Sistema")
            Me.TextBox1.Text = "0"
            Me.Button2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox2.Text) = 0 Then
            MsgBox("Debe seleccionar la Descripción del Administrador de Recursos", vbInformation, "Mensaje del Sistema")
            Me.Button47.Focus()
            Exit Sub
        End If
        Call GRABAR_REGISTRO()
        c46ID_ADMIN = Me.TextBox1.Text
        CERRAR = False
        Me.Close()
    End Sub
    Private Sub GRABAR_REGISTRO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [CAT DE ADMINISTRADORES] (ID_ADMIN, DESCRIPCION, ACTIVO)" &
                                  "values (" & CInt(Me.TextBox1.Text) & ", '" & Trim(Me.TextBox2.Text) & "', '" & Me.CheckBox1.Checked & "')"
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