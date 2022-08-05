Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Situacion_Personal_Ajustes
    Private Sub Frm_Situacion_Personal_Ajustes_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        RemoveHandler Me.DateTimePicker1.ValueChanged, AddressOf DateTimePicker1_ValueChanged
        RemoveHandler Me.DateTimePicker2.ValueChanged, AddressOf DateTimePicker2_ValueChanged
        Me.DateTimePicker1.Value = Frm_Situacion_Personal.DateTimePicker1.Value
        Me.DateTimePicker2.Value = Frm_Situacion_Personal.DateTimePicker2.Value
        Me.DateTimePicker3.Value = Frm_Situacion_Personal.DateTimePicker2.Value
        Call DETALLES()
        AddHandler Me.DateTimePicker1.ValueChanged, AddressOf DateTimePicker1_ValueChanged
        AddHandler Me.DateTimePicker2.ValueChanged, AddressOf DateTimePicker2_ValueChanged
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
    Private Sub Frm_Situacion_Personal_Ajustes_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.Close()
        End If
    End Sub
    Dim D As Double
    Private Sub DETALLES()
        If Me.DateTimePicker1.Value < Me.DateTimePicker2.Value Then
            Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker1.Value, 1, 10) + "', 105), 23)"
            Dim F2 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker2.Value, 1, 10) + "', 105), 23)"
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Dim MiDataTable3 As New DataTable
            Dim MiDataAdapter3 As SqlDataAdapter
            Try
                Conexion.ABD_RRHH()
                MiDataAdapter3 = New SqlDataAdapter("SELECT SUM(VALOR_DIA) FROM [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE ([CODIGO] = '" & Frm_Situacion_Personal.TextBox2.Text & "') AND (DIA BETWEEN " & F1 & " AND " & F2 & ")", Conexion.CxRRHH)
                MiDataTable3.Clear()
                MiDataAdapter3.Fill(MiDataTable3)
                If MiDataTable3.Rows.Count <> 0 Then
                    'Me.TextBox3.Text = Format(Val(MiDataTable3.Rows(0).Item(0).ToString), "##,##0.00") 'YARID
                    Me.TextBox3.Text = Val(MiDataTable3.Rows(0).Item(0).ToString)
                Else
                    Me.TextBox3.Text = "0"
                End If
            Catch ex As Exception
                MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
            Call DIAS_DESCONTADOS()
            Call DIAS_ACREDITADOS()
            Call DIAS_SUBSIDIOS()
            Call DIAS_AI()
            Call DIAS_V()
            Call DIAS_AMP()
            Call DIAS_AMN()
            Call DIAS_OM()
            Call ANTIGUEDAD_PAME()
            Me.TextBox2.Text = Val(Me.TextBox3.Text) + Val(Me.TextBox4.Text) - Val(Me.TextBox1.Text) + Val(Me.TextBox10.Text) - Val(Me.TextBox11.Text)
            'Me.TextBox2.Text = Format(Val(Me.TextBox2.Text), "##,##0.00")   'YARID
            Me.TextBox2.Text = Val(Me.TextBox2.Text)
        End If
    End Sub
    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Call DETALLES()
    End Sub
    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        Call DETALLES()
    End Sub
    Private Sub ANTIGUEDAD_PAME()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable3 As New DataTable
        Dim MiDataAdapter3 As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter3 = New SqlDataAdapter("SELECT CODIGO, ANTIGUEDAD_PAME FROM [dbo].[VISTA MAESTRO DE PERSONAS] WHERE ([CODIGO] = '" & Frm_Situacion_Personal.TextBox2.Text & "')", Conexion.CxRRHH)
            MiDataTable3.Clear()
            MiDataAdapter3.Fill(MiDataTable3)
            If MiDataTable3.Rows.Count <> 0 Then
                'Me.TextBox9.Text = Format(Val(MiDataTable3.Rows(0).Item(1).ToString), "##,##0.00")  'YARID
                Me.TextBox9.Text = Val(MiDataTable3.Rows(0).Item(1).ToString)
            Else
                Me.TextBox9.Text = "0"
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub DIAS_OM()
        Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker1.Value, 1, 10) + "', 105), 23)"
        Dim F2 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker2.Value, 1, 10) + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable3 As New DataTable
        Dim MiDataAdapter3 As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter3 = New SqlDataAdapter("SELECT SUM(VALOR_SIT) FROM [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE ([CODIGO] = '" & Frm_Situacion_Personal.TextBox2.Text & "') AND (ID_SIT_P = 20) AND (DIA BETWEEN " & F1 & " AND " & F2 & ")", Conexion.CxRRHH)
            MiDataTable3.Clear()
            MiDataAdapter3.Fill(MiDataTable3)
            If MiDataTable3.Rows.Count <> 0 Then
                'Me.TextBox7.Text = Format(Math.Abs(Val(MiDataTable3.Rows(0).Item(0).ToString)), "##,##0.00")   'YARID
                Me.TextBox7.Text = Math.Abs(Val(MiDataTable3.Rows(0).Item(0).ToString))
            Else
                Me.TextBox7.Text = "0"
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub DIAS_AMP()
        Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker1.Value, 1, 10) + "', 105), 23)"
        Dim F2 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker2.Value, 1, 10) + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable3 As New DataTable
        Dim MiDataAdapter3 As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter3 = New SqlDataAdapter("SELECT SUM(VALOR_SIT) FROM [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE ([CODIGO] = '" & Frm_Situacion_Personal.TextBox2.Text & "') AND (ID_SIT_P = 38) AND (DIA BETWEEN " & F1 & " AND " & F2 & ")", Conexion.CxRRHH)
            MiDataTable3.Clear()
            MiDataAdapter3.Fill(MiDataTable3)
            If MiDataTable3.Rows.Count <> 0 Then
                Me.TextBox10.Text = Format(Math.Abs(Val(MiDataTable3.Rows(0).Item(0).ToString)), "##,##0.00")   'YARID
                Me.TextBox10.Text = Math.Abs(Val(MiDataTable3.Rows(0).Item(0).ToString))
            Else
                Me.TextBox10.Text = "0"
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub DIAS_AMN()
        Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker1.Value, 1, 10) + "', 105), 23)"
        Dim F2 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker2.Value, 1, 10) + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable3 As New DataTable
        Dim MiDataAdapter3 As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter3 = New SqlDataAdapter("SELECT SUM(VALOR_SIT) FROM [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE ([CODIGO] = '" & Frm_Situacion_Personal.TextBox2.Text & "') AND (ID_SIT_P = 39) AND (DIA BETWEEN " & F1 & " AND " & F2 & ")", Conexion.CxRRHH)
            MiDataTable3.Clear()
            MiDataAdapter3.Fill(MiDataTable3)
            If MiDataTable3.Rows.Count <> 0 Then
                Me.TextBox11.Text = Format(Math.Abs(Val(MiDataTable3.Rows(0).Item(0).ToString)), "##,##0.00")   'YARID
                Me.TextBox11.Text = Math.Abs(Val(MiDataTable3.Rows(0).Item(0).ToString))
            Else
                Me.TextBox11.Text = "0"
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub DIAS_V()
        Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker1.Value, 1, 10) + "', 105), 23)"
        Dim F2 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker2.Value, 1, 10) + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable3 As New DataTable
        Dim MiDataAdapter3 As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter3 = New SqlDataAdapter("SELECT SUM(VALOR_SIT) FROM [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE ([CODIGO] = '" & Frm_Situacion_Personal.TextBox2.Text & "') AND (ID_SIT_P = 17) AND (DIA BETWEEN " & F1 & " AND " & F2 & ")", Conexion.CxRRHH)
            MiDataTable3.Clear()
            MiDataAdapter3.Fill(MiDataTable3)
            If MiDataTable3.Rows.Count <> 0 Then
                Me.TextBox6.Text = Format(Math.Abs(Val(MiDataTable3.Rows(0).Item(0).ToString)), "##,##0.00")    'YARID
                Me.TextBox6.Text = Math.Abs(Val(MiDataTable3.Rows(0).Item(0).ToString))
            Else
                Me.TextBox6.Text = "0"
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub DIAS_AI()
        Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker1.Value, 1, 10) + "', 105), 23)"
        Dim F2 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker2.Value, 1, 10) + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable3 As New DataTable
        Dim MiDataAdapter3 As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter3 = New SqlDataAdapter("SELECT SUM(VALOR_SIT) FROM [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE ([CODIGO] = '" & Frm_Situacion_Personal.TextBox2.Text & "') AND (ID_SIT_P = 3) AND (DIA BETWEEN " & F1 & " AND " & F2 & ")", Conexion.CxRRHH)
            MiDataTable3.Clear()
            MiDataAdapter3.Fill(MiDataTable3)
            If MiDataTable3.Rows.Count <> 0 Then
                'Me.TextBox5.Text = Format(Math.Abs(Val(MiDataTable3.Rows(0).Item(0).ToString)), "##,##0.00")   'YARID
                Me.TextBox5.Text = Math.Abs(Val(MiDataTable3.Rows(0).Item(0).ToString))
            Else
                Me.TextBox5.Text = "0"
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub DIAS_SUBSIDIOS()
        Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker1.Value, 1, 10) + "', 105), 23)"
        Dim F2 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker2.Value, 1, 10) + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable3 As New DataTable
        Dim MiDataAdapter3 As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter3 = New SqlDataAdapter("SELECT COUNT(VALOR_SIT) FROM [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE ([CODIGO] = '" & Frm_Situacion_Personal.TextBox2.Text & "') AND (ID_SIT_P = 14) AND (DIA BETWEEN " & F1 & " AND " & F2 & ")", Conexion.CxRRHH)
            MiDataTable3.Clear()
            MiDataAdapter3.Fill(MiDataTable3)
            If MiDataTable3.Rows.Count <> 0 Then
                'Me.TextBox8.Text = Format(Math.Abs(Val(MiDataTable3.Rows(0).Item(0).ToString)), "##,##0.00")   'YARID
                Me.TextBox8.Text = Math.Abs(Val(MiDataTable3.Rows(0).Item(0).ToString))
            Else
                Me.TextBox8.Text = "0"
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub DIAS_ACREDITADOS()
        Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker1.Value, 1, 10) + "', 105), 23)"
        Dim F2 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker2.Value, 1, 10) + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable3 As New DataTable
        Dim MiDataAdapter3 As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter3 = New SqlDataAdapter("SELECT SUM(VALOR_SIT) FROM [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE ([CODIGO] = '" & Frm_Situacion_Personal.TextBox2.Text & "') AND (ID_SIT_P = 32) AND (DIA BETWEEN " & F1 & " AND " & F2 & ")", Conexion.CxRRHH)
            MiDataTable3.Clear()
            MiDataAdapter3.Fill(MiDataTable3)
            If MiDataTable3.Rows.Count <> 0 Then
                'Me.TextBox4.Text = Format(Math.Abs(Val(MiDataTable3.Rows(0).Item(0).ToString)), "##,##0.00")   'YARID
                Me.TextBox4.Text = Math.Abs(Val(MiDataTable3.Rows(0).Item(0).ToString))
            Else
                Me.TextBox4.Text = "0"
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub DIAS_DESCONTADOS()
        Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker1.Value, 1, 10) + "', 105), 23)"
        Dim F2 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker2.Value, 1, 10) + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable3 As New DataTable
        Dim MiDataAdapter3 As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter3 = New SqlDataAdapter("SELECT SUM(VALOR_SIT) FROM [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE ([CODIGO] = '" & Frm_Situacion_Personal.TextBox2.Text & "') AND (ID_SIT_P = 20 OR ID_SIT_P = 17 OR ID_SIT_P = 3) AND (DIA BETWEEN " & F1 & " AND " & F2 & ")", Conexion.CxRRHH)
            MiDataTable3.Clear()
            MiDataAdapter3.Fill(MiDataTable3)
            If MiDataTable3.Rows.Count <> 0 Then
                Me.TextBox1.Text = Math.Abs(Val(MiDataTable3.Rows(0).Item(0).ToString))
                'Me.TextBox1.Text = Val(Me.TextBox1.Text) + Val(Me.TextBox10.Text) - Val(Me.TextBox11.Text) 'YARID
            Else
                Me.TextBox1.Text = "0"
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