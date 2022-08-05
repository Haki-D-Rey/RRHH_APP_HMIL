Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Configuraciones
    Private Sub Frm_Configuraciones_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Configuraciones_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Show()
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Frm_Config_01.ShowDialog()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Frm_Config_02.ShowDialog()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        MENSAJE = MsgBox("¿Esta seguro de Aplicar este Proceso?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            Call ACTUALIZAR_CLASIFICACION()
            MsgBox("Proceso Finalizado", vbInformation, "Mensaje del Sistema")
        End If
    End Sub
    Dim dID_M_P As Integer
    Dim dSEXO As Integer
    Private Sub ACTUALIZAR_CLASIFICACION()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [VISTA MAESTRO DE CARGOS SEXO] WHERE ID_SITUACION = 1 ORDER BY ID_M_C", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                For Ix = 0 To MiDataTable.Rows.Count - 1
                    dID_M_P = MiDataTable.Rows(Ix).Item(39).ToString
                    dSEXO = MiDataTable.Rows(Ix).Item(50).ToString
                    Call BUSCAR_HIJOS()
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
    Dim TIENE_HIJOS As Boolean
    Private Sub BUSCAR_HIJOS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [MAESTRO DE NUCLEO FAMILIAR] WHERE ID_M_P = " & dID_M_P & " AND ID_PARENTELA > 4 ORDER BY ID_NF", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                TIENE_HIJOS = True
                If dSEXO = 1 Then
                    TIPO_CLAS = 1
                End If
                If dSEXO = 2 Then
                    TIPO_CLAS = 2
                End If
                Call BUSCAR_CLASIFICACION
            Else
                TIENE_HIJOS = False
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim TIPO_CLAS As Integer
    Private Sub BUSCAR_CLASIFICACION()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [MAESTRO DE CLASIFICACION DE PERSONAL] WHERE ID_M_P = " & dID_M_P & " AND ID_C_PE = " & TIPO_CLAS & " ORDER BY ID_C_PERSONAL", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count = 0 Then
                Call AGREGAR_REGISTRO
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub AGREGAR_REGISTRO()
        Call GENERAR_CODIGO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE CLASIFICACION DE PERSONAL] (ID_C_PERSONAL, ID_M_P, ID_C_PE, OBSERVACIONES)" &
                                  "values (" & CInt(CR) & ", " & dID_M_P & ", " & TIPO_CLAS & ", 'INGRESO AUTOMATICO')"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim CR As Integer
    Private Sub GENERAR_CODIGO()
        CR = 0
        CODIGO = 0
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            MiDataAdapter = New SqlDataAdapter("SELECT ID_C_PERSONAL FROM [MAESTRO DE CLASIFICACION DE PERSONAL] ORDER BY ID_C_PERSONAL", Conexion.CxRRHH)
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
            CR = CODIGO
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
End Class