Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient

Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class Frm_Expedientes_Imprimir
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Me.ComboBox2.Focus()
        Me.Close()
    End Sub
    Private Sub Frm_Expedientes_Imprimir_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = ""
            Me.TextBox2.Text = ""
            Me.TextBox3.Text = ""
            Me.TextBox4.Text = ""
            Me.TextBox5.Text = ""
            Me.ComboBox2.Focus()
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Expedientes_Imprimir_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = Frm_Expedientes_Activos.ComboBox12.Text & " " & Frm_Expedientes_Activos.ComboBox6.Text
        Me.TextBox2.Text = Frm_Expedientes_Activos.TextBox30.Text & " " & Frm_Expedientes_Activos.TextBox6.Text
        Me.TextBox3.Text = Frm_Expedientes_Activos.ComboBox4.Text
        Me.TextBox4.Text = "Código N°. " & Frm_Expedientes_Activos.TextBox5.Text
        Me.TextBox5.Text = "Cédula N°. " & Frm_Expedientes_Activos.TextBox10.Text
        Call CARGAR_COMBO_CAT_DE_FORMATOS_DE_SALIDA()
        Call CARGAR_COMBO_CAT_DE_CARNET()
        Me.ComboBox2.Text = "EXPEDIENTE"
        Me.ComboBox2.Focus()
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
                With Me.ComboBox3
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_F_S"
                End With
            Else
                Me.ComboBox3.DataSource = Nothing
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_CARNET()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE CARNET] where ACTIVO = 1 order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_CARNET"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        SELECCION = ""
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        If Me.ComboBox2.Text = "EXPEDIENTE" Then
            If Me.ComboBox3.Text = "CRYSTAL" Then
                PARAMETRO = 4   'EXPEDIENTE DE EMPLEADO
                Frm_Visor_Reportes.CrystalR.ShowExportButton = True
                Frm_Visor_Reportes.CrystalR.ShowPrintButton = True
                SELECCION = "{MAESTRO_DE_PERSONAS.ID_M_P} = " & Frm_Expedientes_Activos.TextBox35.Text
                Frm_Visor_Reportes.ShowDialog()
            End If
            If Me.ComboBox3.Text = "OFFICE (EXCEL\WORD)" Then
                'NO DISPONIBLE
            End If
            If Me.ComboBox3.Text = "PDF" Then
                SELECCION = "{MAESTRO_DE_PERSONAS.ID_M_P} = " & Frm_Expedientes_Activos.TextBox35.Text
                Dim cryRpt As New ReportDocument
                Dim PathRPT As String = "\\Hmil-spc-apprd\e$\v3 SPYC\Reportes\"
                cryRpt.Load(PathRPT & "SPYC004.rpt")
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
        End If
        If Me.ComboBox2.Text = "CARNET" Then
            SELECCION = "{MAESTRO_DE_PERSONAS.ID_M_P} = " & Frm_Expedientes_Activos.TextBox35.Text
            If Me.ComboBox3.Text = "CRYSTAL" Then
                If Me.ComboBox1.Text = "ADMINISTRATIVO" Then
                    PARAMETRO = 5   'CARNET DE IDENTIFICACION DE EMPLEADO (administrativo)
                End If
                If Me.ComboBox1.Text = "TECNICO" Then
                    PARAMETRO = 6   'CARNET DE IDENTIFICACION DE EMPLEADO (personal tecnico)
                End If
                If Me.ComboBox1.Text = "MEDICOS Y MILITARES" Then
                    PARAMETRO = 7   'CARNET DE IDENTIFICACION DE EMPLEADO (medicos y militares)
                End If
                If Me.ComboBox1.Text = "ENFERMERIA" Then
                    PARAMETRO = 8   'CARNET DE IDENTIFICACION DE EMPLEADO (enfermeria)
                End If
                If Me.ComboBox1.Text = "VOLUNTARIADO" Then
                    PARAMETRO = 9   'CARNET DE IDENTIFICACION DE EMPLEADO (voluntariado)
                End If
                If Me.ComboBox1.Text = "POLICLINICO" Then
                    PARAMETRO = 10   'CARNET DE IDENTIFICACION DE EMPLEADO (policlinico)
                End If
                VALOR01 = Me.TextBox1.Text
                VALOR02 = Me.TextBox2.Text
                VALOR03 = Me.TextBox3.Text
                VALOR04 = Me.TextBox4.Text
                VALOR05 = Me.TextBox5.Text

                Frm_Visor_Reportes.CrystalR.ShowExportButton = True
                Frm_Visor_Reportes.CrystalR.ShowPrintButton = True
                Frm_Visor_Reportes.ShowDialog()
            End If
            '/////////////////////////////////////////////////////////////////////////////////////////////////
            '/////////////////////////////////////////////////////////////////////////////////////////////////
            '/////////////////////////////////////////////////////////////////////////////////////////////////
            If Me.ComboBox3.Text = "EXCEL" Then
                'NO DISPONIBLE
            End If
            '/////////////////////////////////////////////////////////////////////////////////////////////////
            '/////////////////////////////////////////////////////////////////////////////////////////////////
            '/////////////////////////////////////////////////////////////////////////////////////////////////
            If Me.ComboBox3.Text = "PDF" Then
                Dim cryRpt As New ReportDocument
                Dim PathRPT As String = "\\Hmil-spc-apprd\e$\v3 SPYC\Reportes\"
                VALOR01 = Me.TextBox1.Text
                VALOR02 = Me.TextBox2.Text
                VALOR03 = Me.TextBox3.Text
                VALOR04 = Me.TextBox4.Text
                VALOR05 = Me.TextBox5.Text
                If Me.ComboBox1.Text = "ADMINISTRATIVO" Then
                    cryRpt.Load(PathRPT & "SPYC005.rpt")
                End If
                If Me.ComboBox1.Text = "TECNICO" Then
                    cryRpt.Load(PathRPT & "SPYC006.rpt")
                End If
                If Me.ComboBox1.Text = "MEDICOS Y MILITARES" Then
                    cryRpt.Load(PathRPT & "SPYC007.rpt")
                End If
                If Me.ComboBox1.Text = "ENFERMERIA" Then
                    cryRpt.Load(PathRPT & "SPYC008.rpt")
                End If
                If Me.ComboBox1.Text = "VOLUNTARIADO" Then
                    cryRpt.Load(PathRPT & "SPYC009.rpt")
                End If
                If Me.ComboBox1.Text = "POLICLINICO" Then
                    cryRpt.Load(PathRPT & "SPYC010.rpt")
                End If
                cryRpt.SetDatabaseLogon("sa", "P@$$W0RD")
                Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
                Dim Ds1 As New CrystalDecisions.Shared.ParameterDiscreteValue() : Ds1.Value = SELECCION_PARAMETRO
                Dim Ds2 As New CrystalDecisions.Shared.ParameterDiscreteValue() : Ds2.Value = USUARIO_IMPRIME
                Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = VALOR01   'NIVEL PROFESIONAL
                Dim V02 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V02.Value = VALOR02   'NOMBRES Y APELLIDOS
                Dim V03 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V03.Value = VALOR03   'CARGO
                Dim V04 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V04.Value = VALOR04   'CODIGO
                Dim V05 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V05.Value = VALOR05   'CEDULA
                RpDatos.Add(V01) : cryRpt.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
                RpDatos.Add(V02) : cryRpt.DataDefinition.ParameterFields("P02").ApplyCurrentValues(RpDatos)
                RpDatos.Add(V03) : cryRpt.DataDefinition.ParameterFields("P03").ApplyCurrentValues(RpDatos)
                RpDatos.Add(V04) : cryRpt.DataDefinition.ParameterFields("P04").ApplyCurrentValues(RpDatos)
                RpDatos.Add(V05) : cryRpt.DataDefinition.ParameterFields("P05").ApplyCurrentValues(RpDatos)
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
        End If
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        If Me.ComboBox2.Text = "LISTA DE MEDALLAS" Then
            Call BUSCAR_USUARIO_A_IMPRIMIR()
            If Me.ComboBox3.Text = "CRYSTAL" Then
                Dim cryRpt As New ReportDocument
                Dim PathRPT As String = "\\Hmil-spc-apprd\e$\v3 SPYC\Reportes\"
                cryRpt.Load(PathRPT & "SPYCE115.rpt")
                cryRpt.SetDatabaseLogon("sa", "P@$$W0RD")
                PARAMETRO = 155
                Frm_Visor_Reportes.CrystalR.ShowExportButton = True
                Frm_Visor_Reportes.CrystalR.ShowPrintButton = True
                SELECCION = "{MAESTRO_DE_PERSONAS.ID_M_P} = " & IDU
                VALOR01 = Me.TextBox1.Text & " " & Me.TextBox2.Text
                Frm_Visor_Reportes.ShowDialog()
            End If
            If Me.ComboBox3.Text = "OFFICE (EXCEL\WORD)" Then
                'NO DISPONIBLE
            End If
            If Me.ComboBox3.Text = "PDF" Then
                SELECCION = "{MAESTRO_DE_PERSONAS.ID_M_P} = " & Frm_Expedientes_Activos.TextBox35.Text
                Dim cryRpt As New ReportDocument
                Dim PathRPT As String = "\\Hmil-spc-apprd\e$\v3 SPYC\Reportes\"
                cryRpt.Load(PathRPT & "SPYCE115.rpt")
                cryRpt.SetDatabaseLogon("sa", "P@$$W0RD")
                VALOR01 = Me.TextBox1.Text & " " & Me.TextBox2.Text
                Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
                Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = VALOR01
                RpDatos.Add(V01) : cryRpt.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
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
        End If
    End Sub
    Dim IDU As Integer
    Private Sub BUSCAR_USUARIO_A_IMPRIMIR()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [dbo].[VISTA MAESTRO DE PERSONAS] WHERE CODIGO = '" & Frm_Expedientes_Activos.TextBox2.Text & "'", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                IDU = MiDataTable.Rows(0).Item(0).ToString
            Else
                IDU = 0
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
End Class