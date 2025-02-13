Imports System
Imports System.Data.SqlClient
Module Conexion
    Public CxRRHH As New SqlConnection              'BASE DE DATOS [RRHH]   ***
    Public CxSIE As New SqlConnection               'BASE DE DATOS [SIE]
    Public CxSUBS As New SqlConnection              'BASE DE DATOS [SUBSIDIOS]
    Public CxTECNOLOGIA As New SqlConnection              'BASE DE DATOS [TECNOLOGIA] 
    '*********************************************************************************************************************************************************
    'BASE DE DATOS [TECNOLOGIA]
    '*********************************************************************************************************************************************************
    Public Sub ABD_TEC()
        Try
            CxTECNOLOGIA.ConnectionString = "Data Source = 127.0.0.1; initial catalog = RRHH; user id = sa; password = Maleisho31102019"
            CxTECNOLOGIA.Open()
        Catch e As Exception
            MsgBox("Error: " & e.Message, vbInformation, "Mensaje del Sistema")
        End Try
    End Sub
    Public Sub CBD_TEC()
        CxTECNOLOGIA.Close()
    End Sub
    '********************************************************************************************************************************
    'BASE DE DATOS [RRHH]
    '********************************************************************************************************************************
    Public Sub ABD_RRHH()
        Try
            CxRRHH.ConnectionString = "Data Source = 127.0.0.1; initial catalog = RRHH; user id = sa; password = Maleisho31102019;"
            CxRRHH.Open()
        Catch e As Exception
            MsgBox("Error: " & e.Message, vbInformation, "Mensaje del Sistema")
        End Try
    End Sub
    Public Sub CBD_RRHH()
        CxRRHH.Close()
    End Sub
    '*********************************************************************************************************************************************************
    'BASE DE DATOS [SIE]
    '*********************************************************************************************************************************************************
    Public Sub ABD_SIE()
        Try
            CxSIE.ConnectionString = "Data Source = HMIL-GRM-APPPRD; initial catalog = SIE; user id = sa; password = P@$$W0RD"
            CxSIE.Open()
        Catch e As Exception
            MsgBox("Error: " & e.Message, vbInformation, "Mensaje del Sistema")
        End Try
    End Sub
    Public Sub CBD_SIE()
        CxSIE.Close()
    End Sub
    '*********************************************************************************************************************************************************
    'BASE DE DATOS [SUBSIDIOS]
    '*********************************************************************************************************************************************************
    Public Sub ABD_SUBS()
        Try
            CxSUBS.ConnectionString = "Data Source = HMIL-DAM-APPPRD; initial catalog = SUBSIDIOS; user id = sa; password = P@$$W0RD"
            CxSUBS.Open()
        Catch e As Exception
            MsgBox("Error: " & e.Message, vbInformation, "Mensaje del Sistema")
        End Try
    End Sub
    Public Sub CBD_SUBS()
        CxSUBS.Close()
    End Sub
End Module
