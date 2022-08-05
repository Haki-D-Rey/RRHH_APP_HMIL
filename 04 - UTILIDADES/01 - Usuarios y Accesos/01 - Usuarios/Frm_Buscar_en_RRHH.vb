Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Buscar_en_RRHH
    Private Sub Frm_Buscar_en_RRHH_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            TextBox1.Focus()
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Buscar_en_RRHH_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = ""
        Me.TextBox1.Focus()
    End Sub
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If (e.KeyCode = Keys.Enter) Then
            cEMPLEADO = CInt(Val(TextBox1.Text))
            TextBox1.Text = Format(CInt(cEMPLEADO), "0000000000")
            If Len(Me.TextBox1.Text) = 0 Or (Me.TextBox1.Text = "0000000000") Then
                MsgBox("El código que se digitó no es válido", vbInformation, "Mensaje del Sistema")
                Me.TextBox1.Focus()
                Exit Sub
            Else
                Call BUSCAR_DATO_EN_RRHH()
                If DATO_EN_RRHH = True Then
                    Me.Close()
                Else
                    MsgBox("El código que se digitó no es válido", vbInformation, "Mensaje del Sistema")
                    Me.TextBox1.Focus()
                    Exit Sub
                End If
            End If
        End If
    End Sub
    Private Sub BUSCAR_1N1A()
        Dim N, A As String
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT APELLIDOS, NOMBRES, CODIGO, ID_ESTADO_P FROM [dbo].[VISTA MAESTRO DE PERSONAS] WHERE CODIGO = '" & bCODIGO & "' AND ID_ESTADO_P = 1 ORDER BY CODIGO", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                N = Trim(Mid(MiDataTable.Rows(0).Item(1).ToString, 1, InStr(MiDataTable.Rows(0).Item(1).ToString, " ")))
                A = Trim(Mid(MiDataTable.Rows(0).Item(0).ToString, 1, InStr(MiDataTable.Rows(0).Item(0).ToString, " ")))
                bCUENTA = N & "." & A
            Else
                N = "."
                A = "."
                bCUENTA = "."
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub BUSCAR_DATO_EN_RRHH()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT CODIGO, NA, D_CARGO_ES, ID_SITUACION, D_N1, D_N2, D_N3, D_N4, D_N5, D_N6, D_N7, D_N8 FROM [dbo].[VISTA MAESTRO DE CARGOS SIT] WHERE CODIGO = '" & Me.TextBox1.Text & "' AND ID_SITUACION = 1 ORDER BY CODIGO", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                bCODIGO = MiDataTable.Rows(0).Item(0).ToString
                bEMPLEADO = MiDataTable.Rows(0).Item(1).ToString
                bCARGO = MiDataTable.Rows(0).Item(2).ToString
                If Not IsDBNull(MiDataTable.Rows(0).Item(4).ToString) Then : bD_N1 = MiDataTable.Rows(0).Item(4).ToString : Else : bD_N1 = "" : End If
                If Not IsDBNull(MiDataTable.Rows(0).Item(5).ToString) Then : bD_N2 = MiDataTable.Rows(0).Item(5).ToString : Else : bD_N2 = "" : End If
                If Not IsDBNull(MiDataTable.Rows(0).Item(6).ToString) Then : bD_N3 = MiDataTable.Rows(0).Item(6).ToString : Else : bD_N3 = "" : End If
                If Not IsDBNull(MiDataTable.Rows(0).Item(7).ToString) Then : bD_N4 = MiDataTable.Rows(0).Item(7).ToString : Else : bD_N4 = "" : End If
                If Not IsDBNull(MiDataTable.Rows(0).Item(8).ToString) Then : bD_N5 = MiDataTable.Rows(0).Item(8).ToString : Else : bD_N5 = "" : End If
                If Not IsDBNull(MiDataTable.Rows(0).Item(9).ToString) Then : bD_N6 = MiDataTable.Rows(0).Item(9).ToString : Else : bD_N6 = "" : End If
                If Not IsDBNull(MiDataTable.Rows(0).Item(10).ToString) Then : bD_N7 = MiDataTable.Rows(0).Item(10).ToString : Else : bD_N7 = "" : End If
                If Not IsDBNull(MiDataTable.Rows(0).Item(11).ToString) Then : bD_N8 = MiDataTable.Rows(0).Item(11).ToString : Else : bD_N8 = "" : End If
                Call BUSCAR_1N1A()
                DATO_EN_RRHH = True
            Else
                bCODIGO = "."
                bEMPLEADO_A = "."
                bCARGO = "."
                bD_N1 = "."
                bD_N2 = "."
                bD_N3 = "."
                bD_N4 = "."
                bD_N5 = "."
                bD_N6 = "."
                bD_N7 = "."
                bD_N8 = "."
                DATO_EN_RRHH = False
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
End Class