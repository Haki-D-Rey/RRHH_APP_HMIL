Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Public Class Frm_Expedientes_Calcular_Fechas
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub Frm_Expedientes_Calcular_Fechas_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Expedientes_Calcular_Fechas_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call BORRAR_CONTRATOS()
        Call AGREGAR_REGISTRO_DE_CONTRATOS()
        CERRAR = False
        Me.Close()
    End Sub
    Private Sub AGREGAR_REGISTRO_CYE()
        Dim fFECHA_I As String
        If ceFI <> "" Then
            fFECHA_I = Mid(ceFI, 1, 10)
        Else
            fFECHA_I = "NULL"
        End If
        Dim f1 As Object = If(fFECHA_I <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_I + "', 105), 23)", "NULL")

        Dim fFECHA_F As String
        If ceFF <> "" Then
            fFECHA_F = Mid(ceFF, 1, 10)
        Else
            fFECHA_F = "NULL"
        End If
        Dim f2 As Object = If(fFECHA_F <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_F + "', 105), 23)", "NULL")

        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE CONTRATOS Y EVALUACIONES] (ID_M_CYE, ID_M_P, ID_CONT, FECHA_I, FECHA_F, FIRMADO, ID_C_E, OBSERVACIONES)" &
                                  "values (" & CInt(vID_M_CYE) & ", " & Frm_Expedientes_Activos.TextBox35.Text & ", " & tipo & ", " & f1 & ", " & f2 & ", '" & firma & "', 1, 'INGRESO AUTOMATICO DESDE EL MODULO DE CONTRATOS')"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim vID_M_CYE As Integer
    Private Sub GENERAR_CODIGO_CYE()
        CODIGO = 0
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            MiDataAdapter = New SqlDataAdapter("SELECT ID_M_CYE FROM [MAESTRO DE CONTRATOS Y EVALUACIONES] ORDER BY ID_M_CYE", Conexion.CxRRHH)
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
            vID_M_CYE = CODIGO
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim ceFI As String
    Dim ceFF As String
    Dim tipo As Integer
    Dim firma As Boolean

    Dim ceI1CONT, ceF1CONT, ceI1EVAL, ceF1EVAL, ceI2CONT, ceF2CONT, ceI2EVAL, ceF2EVAL, ceF3CONT As Object
    Private Sub AGREGAR_REGISTRO_DE_CONTRATOS()
        '1 CONTRATO
        ceI1CONT = Mid(Frm_Expedientes_Activos.DateTimePicker2.Value, 1, 10)     'IGUAL A LA FECHA DE INGRESO
        ceF1CONT = Mid(Frm_Expedientes_Activos.DateTimePicker2.Value.AddMonths(3), 1, 10)
        ceF1CONT = CDate(ceF1CONT).AddDays(-1)              '3 MESES + 1 DIA DESPUES DE LA FECHA DE INGRESO
        tipo = 1
        ceFI = ceI1CONT
        ceFF = ceF1CONT
        firma = True
        Call GENERAR_CODIGO_CYE()
        Call AGREGAR_REGISTRO_CYE()
        '1 EVALUACION
        ceI1EVAL = Mid(Frm_Expedientes_Activos.DateTimePicker2.Value.AddMonths(2), 1, 10)    '2 MESES DESPUES DE LA FECHA DE INGRESO
        ceF1EVAL = CDate(ceI1EVAL).AddMonths(1)                         '1 MES DESPUES DE INICIAR LA 1 EVALUACION
        tipo = 2
        ceFI = ceI1EVAL
        ceFF = ceF1EVAL
        firma = False
        Call GENERAR_CODIGO_CYE()
        Call AGREGAR_REGISTRO_CYE()
        '2 CONTRATO
        ceI2CONT = CDate(ceF1CONT).AddDays(1)               '1 DIA DESPUES DE FINALIZAR EL PRIMER CONTRATO
        ceF2CONT = CDate(ceF1CONT).AddMonths(6)             '6 MESES DESPUES DE FINALIZAR EL PRIMER CONTRATO
        tipo = 3
        ceFI = ceI2CONT
        ceFF = ceF2CONT
        firma = False
        Call GENERAR_CODIGO_CYE()
        Call AGREGAR_REGISTRO_CYE()
        '2 EVALUACION
        ceI2EVAL = CDate(ceI2CONT).AddMonths(4)             '4 MESES DESPUES DE INICIAR LA PRIMERA EVALUACION
        ceF2EVAL = CDate(ceI2EVAL).AddMonths(1)             '1 MES DESPUES DE HABER INICIADO LA SEGUNDA EVALUACION
        tipo = 4
        ceFI = ceI2EVAL
        ceFF = ceF2EVAL
        firma = False
        Call GENERAR_CODIGO_CYE()
        Call AGREGAR_REGISTRO_CYE()
        'CONTRATO INDEFINIDO
        ceF3CONT = CDate(ceF2CONT).AddDays(1)               '1 DIA  DESPUES DE HABER FINALIZADO EL SEGUNDO CONTRATO
        tipo = 5
        ceFI = ceF3CONT
        ceFF = ""
        firma = False
        Call GENERAR_CODIGO_CYE()
        Call AGREGAR_REGISTRO_CYE()
    End Sub
    Private Sub BORRAR_CONTRATOS()
        Me.Cursor = Cursors.WaitCursor
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE CONTRATOS Y EVALUACIONES] WHERE [ID_M_P] = " & Frm_Expedientes_Activos.TextBox35.Text & ""
            MiSqlCommand.ExecuteNonQuery()
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