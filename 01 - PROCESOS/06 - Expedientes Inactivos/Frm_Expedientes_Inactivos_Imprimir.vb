Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient

Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class Frm_Expedientes_Inactivos_Imprimir
    Private Sub Frm_Expedientes_Inactivos_Imprimir_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.ComboBox1.Focus()
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Expedientes_Inactivos_Imprimir_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call CARGAR_COMBO_CAT_DE_FORMATOS_DE_SALIDA()
    End Sub
    Private Sub Button03_Click(sender As Object, e As EventArgs) Handles Button03.Click
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
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE FORMATOS DE SALIDA] WHERE ID_F_S <> 2 ORDER BY ID_F_S", Conexion.CxRRHH)
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
    Private Sub Button01_Click(sender As Object, e As EventArgs) Handles Button01.Click
        SELECCION = ""
        If Me.ComboBox1.Text = "CRYSTAL" Then
            PARAMETRO = 12   'EXPEDIENTE DE EMPLEADO INACTIVO
            Frm_Visor_Reportes.CrystalR.ShowExportButton = True
            Frm_Visor_Reportes.CrystalR.ShowPrintButton = True
            SELECCION = "{MAESTRO_DE_PERSONAS.ID_M_P} = " & Frm_Expedientes_Inactivos.TextBox35.Text
            Frm_Visor_Reportes.ShowDialog()
        End If
        If Me.ComboBox1.Text = "OFFICE (EXCEL\WORD)" Then
            'NO DISPONIBLE
        End If
        If Me.ComboBox1.Text = "PDF" Then
            SELECCION = "{MAESTRO_DE_PERSONAS.ID_M_P} = " & Frm_Expedientes_Inactivos.TextBox35.Text
            Dim cryRpt As New ReportDocument
            Dim PathRPT As String = "\\Hmil-spc-apprd\e$\v3 SPYC\Reportes\"
            cryRpt.Load(PathRPT & "SPYC012.rpt")
            cryRpt.SetDatabaseLogon("sa", "P@$$W0RD")
            If SELECCION_PARAMETRO = "" Then
                SELECCION_PARAMETRO = "*"
            End If
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