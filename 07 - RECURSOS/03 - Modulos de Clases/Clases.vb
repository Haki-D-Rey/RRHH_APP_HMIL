Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Net
Public Class Clases
    Public Shared Sub BUSCAR_FECHA_Y_HORA()
        Dim MiDataTable As New DataTable
        If Conexion.CxTECNOLOGIA.State = ConnectionState.Open Then
            Conexion.CBD_TEC()
        End If
        Try
            Conexion.ABD_TEC()
            MiDataAdapter = New SqlDataAdapter("SELECT CONVERT (datetime, SYSDATETIME())", Conexion.CxTECNOLOGIA)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                FECHA_SERVIDOR = MiDataTable.Rows(0).Item(0).ToString
                HORA_SERVIDOR = MiDataTable.Rows(0).Item(0).ToString
                HORA = HORA_SERVIDOR.Hour & ":" & HORA_SERVIDOR.Minute & ":" & HORA_SERVIDOR.Second
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxTECNOLOGIA.State = ConnectionState.Open Then
                Conexion.CBD_TEC()
            End If
        End Try
    End Sub
    Public Shared Function GetIP() As String
        Dim valorIp As String
        valorIp = Dns.GetHostEntry(My.Computer.Name).AddressList.FirstOrDefault(Function(i) i.AddressFamily = Sockets.AddressFamily.InterNetwork).ToString()
        Return valorIp
    End Function
    'Public Shared Sub INSERTAR_AUDITOR_GENERAL_SPYC(ByVal aPROCESO As String, ByVal aMODULO As String, ByVal aID_REGISTRO As String, ByVal aCODIGO As String, ByVal aFECHA As String, ByVal aHORA As String, ByVal aUSUARIO_EJECUTA As String, ByVal aSESION_USUARIO_ACT As String, ByVal aTERMINAL_ACT As String, ByVal aIP_ACT As String)
    '    If isuID_TIPO_USUARIO > 1 Then
    '        Dim fFECHA_I1 As String = Mid(aFECHA, 1, 10)
    '        Dim F1 As Object = If(fFECHA_I1 <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_I1 + "', 105), 23)", "NULL")
    '        Dim MiDataTable As New DataTable
    '        If Conexion.CxRRHH.State = ConnectionState.Open Then
    '            Conexion.CBD_RRHH()
    '        End If
    '        Conexion.ABD_RRHH()
    '        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
    '        Try
    '            MiSqlCommand.CommandText = "INSERT INTO [AUDITOR GENERAL RRHH] (PROCESO, MODULO, ID_REGISTRO, CODIGO, FECHA, HORA, USUARIO_EJECUTA, SESION_USUARIO_ACT, TERMINAL_ACT, IP_ACT)" &
    '                                      "values ('" & aPROCESO & "', '" & aMODULO & "', '" & aID_REGISTRO & "', '" & aCODIGO & "', " & F1 & ", '" & aHORA & "', '" & aUSUARIO_EJECUTA & "', '" & aSESION_USUARIO_ACT & "', '" & aTERMINAL_ACT & "', '" & aIP_ACT & "')"
    '            MiSqlCommand.ExecuteNonQuery()
    '        Catch ex As Exception
    '            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
    '        Finally
    '            If Conexion.CxRRHH.State = ConnectionState.Open Then
    '                Conexion.CBD_RRHH()
    '            End If
    '        End Try
    '    End If
    'End Sub

End Class
