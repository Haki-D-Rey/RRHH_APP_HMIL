Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Visor_Documentos_Inactivos
    Private Sub Frm_Visor_Documentos_Inactivos_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Visor_Documentos_Inactivos_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        PDFcontendor.src = vmdDIRECTORIO & "\" & vmdNOMBRE
    End Sub
End Class