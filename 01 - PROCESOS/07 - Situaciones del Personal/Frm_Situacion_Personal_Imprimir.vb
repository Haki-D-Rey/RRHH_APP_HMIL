Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient

Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class Frm_Situacion_Personal_Imprimir
    Private Sub Frm_Situacion_Personal_Imprimir_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.ComboBox1.Focus()
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Situacion_Personal_Imprimir_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Call CARGAR_COMBO_CAT_DE_FORMATOS_DE_SALIDA()
        Me.ComboBox1.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.ComboBox1.Focus()
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_FORMATOS_DE_SALIDA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE FORMATOS DE SALIDA] ORDER BY ID_F_S", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                With Me.ComboBox1
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_F_S"
                End With
            Else
                Me.ComboBox1.DataSource = Nothing
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button6_Click(sender As Object, e As System.EventArgs) Handles Button6.Click
        SELECCION = ""
        Dim D1 As String
        D1 = Mid(Frm_Situacion_Personal.DateTimePicker1.Value, 1, 2)
        Dim M1 As String
        M1 = Mid(Frm_Situacion_Personal.DateTimePicker1.Value, 4, 2)
        Dim A1 As String
        A1 = Mid(Frm_Situacion_Personal.DateTimePicker1.Value, 7, 4)
        Dim D2 As String
        D2 = Mid(Frm_Situacion_Personal.DateTimePicker2.Value, 1, 2)
        Dim M2 As String
        M2 = Mid(Frm_Situacion_Personal.DateTimePicker2.Value, 4, 2)
        Dim A2 As String
        A2 = Mid(Frm_Situacion_Personal.DateTimePicker2.Value, 7, 4)
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        If Me.ComboBox1.Text = "CRYSTAL" Then
            SELECCION = "({MAESTRO_DE_PERSONAS.ID_M_P} = " & id_EMPLEADO & ") AND ({MAESTRO_DE_SITUACION_DE_PERSONAL.DIA} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ") TO CDateTime (" & A2 & ", " & M2 & ", " & D2 & "))"
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            Frm_Visor_Reportes.CrystalR.ShowExportButton = True
            Frm_Visor_Reportes.CrystalR.ShowPrintButton = True
            PARAMETRO = 13
            Frm_Visor_Reportes.ShowDialog()
        End If
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        If Me.ComboBox1.Text = "OFFICE (EXCEL\WORD)" Then
            Dim rd = New ReportDocument()
            Dim PathRPT As String = "\\Hmil-spc-apprd\e$\v3 SPYC\Excel\"
            rd.Load(PathRPT & "SPYC013x.rpt")
            SELECCION = "({MAESTRO_DE_PERSONAS.ID_M_P} = " & id_EMPLEADO & ") AND ({MAESTRO_DE_SITUACION_DE_PERSONAL.DIA} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ") TO CDateTime (" & A2 & ", " & M2 & ", " & D2 & "))"
            rd.SetDatabaseLogon("sa", "P@$$W0RD")
            If SELECCION_PARAMETRO = "" Then
                SELECCION_PARAMETRO = "*"
            End If
            On Error GoTo SALIR
            Me.Cursor = Cursors.WaitCursor
            Dim NOMBRE_ARMADO As String
            NOMBRE_ARMADO = "tmpQUERY_" & PARAMETRO & "_" & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00") & ".XLS"
            rd.RecordSelectionFormula = SELECCION
            Dim objExcelOptions As ExcelFormatOptions = New ExcelFormatOptions
            objExcelOptions.ExcelUseConstantColumnWidth = False
            rd.ExportOptions.FormatOptions = objExcelOptions
            Dim strExportFile As String = "C:\WINDOWS\TEMP\" & NOMBRE_ARMADO
            rd.ExportOptions.ExportDestinationType = ExportDestinationType.DiskFile
            rd.ExportOptions.ExportFormatType = ExportFormatType.Excel
            objExcelOptions.ExcelUseConstantColumnWidth = False
            rd.ExportOptions.FormatOptions = objExcelOptions
            Dim objOptions As DiskFileDestinationOptions = New DiskFileDestinationOptions
            objOptions.DiskFileName = strExportFile
            rd.ExportOptions.DestinationOptions = objOptions
            rd.Export()
            objOptions = Nothing
            rd = Nothing
            Process.Start("Excel.exe", "C:\WINDOWS\TEMP\" & NOMBRE_ARMADO)
            Me.Cursor = Cursors.Default
            Exit Sub
SALIR:
            MsgBox("Se ha generado un error, posiblemente no se han seleccionado los parametros requeridos para generar el archivo de Excel", vbInformation, "Mensaje del Sistema")
            Me.Cursor = Cursors.Default
        End If
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        If Me.ComboBox1.Text = "PDF" Then
            Dim cryRpt As New ReportDocument
            Dim PathRPT As String = "\\Hmil-spc-apprd\e$\v3 SPYC\Reportes\"
            SELECCION = "({MAESTRO_DE_PERSONAS.ID_M_P} = " & id_EMPLEADO & ") AND ({MAESTRO_DE_SITUACION_DE_PERSONAL.DIA} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ") TO CDateTime (" & A2 & ", " & M2 & ", " & D2 & "))"
            cryRpt.Load(PathRPT & "SPYC013.rpt")
            cryRpt.SetDatabaseLogon("sa", "P@$$W0RD")
            If SELECCION_PARAMETRO = "" Then
                SELECCION_PARAMETRO = "*"
            End If
            On Error GoTo SALIR
            Me.Cursor = Cursors.WaitCursor
            Dim NOMBRE_ARMADO As String
            NOMBRE_ARMADO = "tmpQUERY_" & PARAMETRO & "_" & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00") & ".pdf"
            Dim CrExportOptions As ExportOptions
            Dim CrDiskFileDestinationOptions As New DiskFileDestinationOptions()
            Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()
            cryRpt.RecordSelectionFormula = SELECCION
            CrDiskFileDestinationOptions.DiskFileName = "C:\WINDOWS\TEMP\" & NOMBRE_ARMADO
            CrExportOptions = cryRpt.ExportOptions
            With CrExportOptions
                .ExportDestinationType = ExportDestinationType.DiskFile
                .ExportFormatType = ExportFormatType.PortableDocFormat
                .DestinationOptions = CrDiskFileDestinationOptions
                .FormatOptions = CrFormatTypeOptions
            End With
            cryRpt.Export()
            Process.Start("C:\WINDOWS\TEMP\" & NOMBRE_ARMADO)
            Me.Cursor = Cursors.Default
            Exit Sub
            Me.Cursor = Cursors.Default
        End If
    End Sub
End Class