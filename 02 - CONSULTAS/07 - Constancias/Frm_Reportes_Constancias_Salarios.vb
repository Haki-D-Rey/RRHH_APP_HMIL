Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Public Class Frm_Reportes_Constancias_Salarios
    Private Sub Frm_Reportes_Constancias_Salarios_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            On Error GoTo SALIR
            If CheckedListBox1.CheckedItems.Count <> 0 Then
                TOTAL_SALARIOS = 0
                For i As Integer = 0 To CheckedListBox1.Items.Count - 1
                    If CheckedListBox1.GetItemChecked(i) Then
                        Dim v As Integer = InStr(Me.CheckedListBox1.Items(i).ToString, "-")
                        v = v + 1
                        TOTAL_SALARIOS = TOTAL_SALARIOS + CDbl(Mid(Me.CheckedListBox1.Items(i).ToString, v, Len(Me.CheckedListBox1.Items(i).ToString)))
                    End If
                Next
            End If
            Me.CheckedListBox1.Items.Clear()
SALIR:
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Reportes_Constancias_Salarios_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.CheckedListBox1.Items.Clear()
        Call CARGAR_LISTA_COLUMNAS_01()
    End Sub
    Private Sub CARGAR_LISTA_COLUMNAS_01()
        Me.Cursor = Cursors.WaitCursor
        I = 0
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [dbo].[VISTA MAESTRO DE NOMINAS Y SALARIOS CONSTANCIAS] WHERE CODIGO = '" & Frm_Reportes_Constancias.TextBox1.Text & "' AND ACTIVO = 1", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                Application.DoEvents()
                For I = 0 To MiDataTable.Rows.Count - 1
                    Me.CheckedListBox1.Items.Add(MiDataTable.Rows(I).Item(6).ToString & " - " & MiDataTable.Rows(I).Item(8).ToString)
                Next
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
        Me.Cursor = Cursors.Default
    End Sub
End Class