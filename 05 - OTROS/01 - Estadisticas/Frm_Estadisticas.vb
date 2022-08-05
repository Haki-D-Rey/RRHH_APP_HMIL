Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Public Class Frm_Estadisticas
    Private Sub Frm_Estadisticas_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            ESTADISTICAS = False
            Frm_Principal.Timer1.Start()
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Estadisticas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Frm_Principal.Timer1.Stop()
        ESTADISTICAS = True
        RemoveHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
        RemoveHandler Me.DateTimePicker1.ValueChanged, AddressOf DateTimePicker1_ValueChanged
        RemoveHandler Me.DateTimePicker2.ValueChanged, AddressOf DateTimePicker2_ValueChanged
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_ESTADISTICAS()
        Call Clases.BUSCAR_FECHA_Y_HORA()
        Me.DateTimePicker1.Value = Mid(FECHA_SERVIDOR, 1, 10)
        Me.DateTimePicker2.Value = Mid(FECHA_SERVIDOR, 1, 10)
        Call SELECCION()
        AddHandler Me.DateTimePicker2.ValueChanged, AddressOf DateTimePicker2_ValueChanged
        AddHandler Me.DateTimePicker1.ValueChanged, AddressOf DateTimePicker1_ValueChanged
        AddHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_ESTADISTICAS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE ESTADISTICAS] ORDER BY ID_T_ESTADISTICA", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_T_ESTADISTICA"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_DATA_8()
        Dim fFECHA_I As String
        fFECHA_I = Mid(Me.DateTimePicker1.Value, 1, 10)
        Dim f1 As Object = If(fFECHA_I <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_I + "', 105), 23)", "NULL")
        Dim fFECHA_I2 As String
        fFECHA_I2 = Mid(Me.DateTimePicker2.Value, 1, 10)
        Dim f2 As Object = If(fFECHA_I2 <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_I2 + "', 105), 23)", "NULL")
        CADENAsql = "SELECT ROW_NUMBER() OVER( ORDER BY [D_N5]) AS ID, [D_N5], COUNT([D_SIT_P]) AS TOTAL FROM [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL 1] WHERE ([DIA] BETWEEN " & f1 & " AND " & f2 & ") AND ([ID_SITUACION] = 1) AND (ID_SIT_P <> 1 AND ID_SIT_P <> 23 AND ID_SIT_P <> 32) GROUP BY [D_N5] ORDER BY ID"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim cmd As New SqlCommand(CADENAsql, Conexion.CxRRHH)
        cmd.CommandType = CommandType.Text
        Dim da As New SqlDataAdapter(cmd)
        Dim dt As New DataTable
        da.Fill(dt)
        Grafico1.Series(0).Points.Clear()
        Grafico1.Series(0).IsValueShownAsLabel = True
        Grafico1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Line
        Grafico1.ChartAreas(0).Area3DStyle.Enable3D = False
        Grafico1.Series(0).Color = Color.RoyalBlue
        Grafico1.ChartAreas(0).AxisX.IntervalAutoMode = 1
        Grafico1.Series(0).XValueMember = "D_N5"
        Grafico1.Series(0).YValueMembers = "TOTAL"
        Grafico1.DataSource = dt
        '-----------------------------------------------------------
        Me.DGV.DataSource = Nothing
        CADENAsql = "SELECT D_N1, D_N2, D_N3, D_N4, D_N5, [D_SIT_P], COUNT([ID_M_P]) AS TOTAL FROM [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL 1] WHERE ([DIA] BETWEEN " & f1 & " AND " & f2 & ") AND ([ID_SITUACION] = 1) AND (ID_SIT_P <> 1 AND ID_SIT_P <> 23 AND ID_SIT_P <> 32) GROUP BY D_N1, D_N2, D_N3, D_N4, D_N5, [D_SIT_P], ORDEN ORDER BY ORDEN"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE SITUACION DE PERSONAL 1]")
        DGV.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE SITUACION DE PERSONAL 1]")
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        '-----------------------------------------------------------

        Me.DGV.DefaultCellStyle.ForeColor = Color.Black

        Me.DGV.Columns(0).HeaderText = "Nivel 1"
        Me.DGV.Columns(0).Width = 40
        Me.DGV.Columns(0).Visible = True
        Me.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(1).HeaderText = "Nivel 2"
        Me.DGV.Columns(1).Width = 40
        Me.DGV.Columns(1).Visible = True
        Me.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(2).HeaderText = "Nivel 3"
        Me.DGV.Columns(2).Width = 80
        Me.DGV.Columns(2).Visible = True
        Me.DGV.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(3).HeaderText = "Nivel 4"
        Me.DGV.Columns(3).Width = 100
        Me.DGV.Columns(3).Visible = True
        Me.DGV.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(4).HeaderText = "Nivel 5"
        Me.DGV.Columns(4).Width = 110
        Me.DGV.Columns(4).Visible = True
        Me.DGV.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(5).HeaderText = "Situacion"
        Me.DGV.Columns(5).Width = 70
        Me.DGV.Columns(5).Visible = True
        Me.DGV.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(6).HeaderText = "Total"
        Me.DGV.Columns(6).Width = 40
        Me.DGV.Columns(6).Visible = True
        Me.DGV.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        'Me.DGV.Columns(7).HeaderText = "Orden"
        'Me.DGV.Columns(7).Width = 0
        'Me.DGV.Columns(7).Visible = False
    End Sub
    Private Sub CARGAR_DATA_7()
        Dim fFECHA_I As String
        fFECHA_I = Mid(Me.DateTimePicker1.Value, 1, 10)
        Dim f1 As Object = If(fFECHA_I <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_I + "', 105), 23)", "NULL")
        Dim fFECHA_I2 As String
        fFECHA_I2 = Mid(Me.DateTimePicker2.Value, 1, 10)
        Dim f2 As Object = If(fFECHA_I2 <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_I2 + "', 105), 23)", "NULL")
        CADENAsql = "SELECT ROW_NUMBER() OVER( ORDER BY [D_N4]) AS ID, [D_N4], COUNT([D_SIT_P]) AS TOTAL FROM [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL 1] WHERE ([DIA] BETWEEN " & f1 & " AND " & f2 & ") AND ([ID_SITUACION] = 1) AND (ID_SIT_P <> 1 AND ID_SIT_P <> 23 AND ID_SIT_P <> 32) GROUP BY [D_N4] ORDER BY ID"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim cmd As New SqlCommand(CADENAsql, Conexion.CxRRHH)
        cmd.CommandType = CommandType.Text
        Dim da As New SqlDataAdapter(cmd)
        Dim dt As New DataTable
        da.Fill(dt)
        Grafico1.Series(0).Points.Clear()
        Grafico1.Series(0).IsValueShownAsLabel = True
        Grafico1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Line
        Grafico1.ChartAreas(0).Area3DStyle.Enable3D = False
        Grafico1.Series(0).Color = Color.RoyalBlue
        Grafico1.ChartAreas(0).AxisX.IntervalAutoMode = 1
        Grafico1.Series(0).XValueMember = "D_N4"
        Grafico1.Series(0).YValueMembers = "TOTAL"
        Grafico1.DataSource = dt
        '-----------------------------------------------------------
        Me.DGV.DataSource = Nothing
        CADENAsql = "SELECT D_N1, D_N2, D_N3, D_N4, [D_SIT_P], COUNT([ID_M_P]) AS TOTAL FROM [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL 1] WHERE ([DIA] BETWEEN " & f1 & " AND " & f2 & ") AND ([ID_SITUACION] = 1) AND (ID_SIT_P <> 1 AND ID_SIT_P <> 23 AND ID_SIT_P <> 32) GROUP BY D_N1, D_N2, D_N3, D_N4, [D_SIT_P], ORDEN ORDER BY ORDEN"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE SITUACION DE PERSONAL 1]")
        DGV.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE SITUACION DE PERSONAL 1]")
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        '-----------------------------------------------------------

        Me.DGV.DefaultCellStyle.ForeColor = Color.Black

        Me.DGV.Columns(0).HeaderText = "Nivel 1"
        Me.DGV.Columns(0).Width = 80
        Me.DGV.Columns(0).Visible = True
        Me.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(1).HeaderText = "Nivel 2"
        Me.DGV.Columns(1).Width = 80
        Me.DGV.Columns(1).Visible = True
        Me.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(2).HeaderText = "Nivel 3"
        Me.DGV.Columns(2).Width = 80
        Me.DGV.Columns(2).Visible = True
        Me.DGV.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(3).HeaderText = "Nivel 4"
        Me.DGV.Columns(3).Width = 100
        Me.DGV.Columns(3).Visible = True
        Me.DGV.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(4).HeaderText = "Situacion"
        Me.DGV.Columns(4).Width = 100
        Me.DGV.Columns(4).Visible = True
        Me.DGV.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(5).HeaderText = "Total"
        Me.DGV.Columns(5).Width = 40
        Me.DGV.Columns(5).Visible = True
        Me.DGV.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        'Me.DGV.Columns(7).HeaderText = "Orden"
        'Me.DGV.Columns(7).Width = 0
        'Me.DGV.Columns(7).Visible = False
    End Sub
    Private Sub CARGAR_DATA_6()
        Dim fFECHA_I As String
        fFECHA_I = Mid(Me.DateTimePicker1.Value, 1, 10)
        Dim f1 As Object = If(fFECHA_I <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_I + "', 105), 23)", "NULL")
        Dim fFECHA_I2 As String
        fFECHA_I2 = Mid(Me.DateTimePicker2.Value, 1, 10)
        Dim f2 As Object = If(fFECHA_I2 <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_I2 + "', 105), 23)", "NULL")
        CADENAsql = "SELECT ROW_NUMBER() OVER( ORDER BY [D_N3]) AS ID, [D_N3], COUNT([D_SIT_P]) AS TOTAL FROM [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL 1] WHERE ([DIA] BETWEEN " & f1 & " AND " & f2 & ") AND ([ID_SITUACION] = 1) AND (ID_SIT_P <> 1 AND ID_SIT_P <> 23 AND ID_SIT_P <> 32) GROUP BY [D_N3] ORDER BY ID"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim cmd As New SqlCommand(CADENAsql, Conexion.CxRRHH)
        cmd.CommandType = CommandType.Text
        Dim da As New SqlDataAdapter(cmd)
        Dim dt As New DataTable
        da.Fill(dt)
        Grafico1.Series(0).Points.Clear()
        Grafico1.Series(0).IsValueShownAsLabel = True
        Grafico1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Line
        Grafico1.ChartAreas(0).Area3DStyle.Enable3D = False
        Grafico1.Series(0).Color = Color.RoyalBlue
        Grafico1.ChartAreas(0).AxisX.IntervalAutoMode = 1
        Grafico1.Series(0).XValueMember = "D_N3"
        Grafico1.Series(0).YValueMembers = "TOTAL"
        Grafico1.DataSource = dt
        '-----------------------------------------------------------
        Me.DGV.DataSource = Nothing
        CADENAsql = "SELECT D_N1, D_N2, D_N3, [D_SIT_P], COUNT([ID_M_P]) AS TOTAL FROM [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL 1] WHERE ([DIA] BETWEEN " & f1 & " AND " & f2 & ") AND ([ID_SITUACION] = 1) AND (ID_SIT_P <> 1 AND ID_SIT_P <> 23 AND ID_SIT_P <> 32) GROUP BY D_N1, D_N2, D_N3, [D_SIT_P], ORDEN ORDER BY ORDEN"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE SITUACION DE PERSONAL 1]")
        DGV.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE SITUACION DE PERSONAL 1]")
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        '-----------------------------------------------------------

        Me.DGV.DefaultCellStyle.ForeColor = Color.Black

        Me.DGV.Columns(0).HeaderText = "Nivel 1"
        Me.DGV.Columns(0).Width = 100
        Me.DGV.Columns(0).Visible = True
        Me.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(1).HeaderText = "Nivel 2"
        Me.DGV.Columns(1).Width = 90
        Me.DGV.Columns(1).Visible = True
        Me.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(2).HeaderText = "Nivel 3"
        Me.DGV.Columns(2).Width = 150
        Me.DGV.Columns(2).Visible = True
        Me.DGV.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(3).HeaderText = "Situacion"
        Me.DGV.Columns(3).Width = 100
        Me.DGV.Columns(3).Visible = True
        Me.DGV.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(4).HeaderText = "Total"
        Me.DGV.Columns(4).Width = 40
        Me.DGV.Columns(4).Visible = True
        Me.DGV.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        'Me.DGV.Columns(7).HeaderText = "Orden"
        'Me.DGV.Columns(7).Width = 0
        'Me.DGV.Columns(7).Visible = False
    End Sub
    Private Sub CARGAR_DATA_5()
        Dim fFECHA_I As String
        fFECHA_I = Mid(Me.DateTimePicker1.Value, 1, 10)
        Dim f1 As Object = If(fFECHA_I <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_I + "', 105), 23)", "NULL")
        Dim fFECHA_I2 As String
        fFECHA_I2 = Mid(Me.DateTimePicker2.Value, 1, 10)
        Dim f2 As Object = If(fFECHA_I2 <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_I2 + "', 105), 23)", "NULL")
        CADENAsql = "SELECT [D_SIT_P], COUNT([ID_M_P]) AS TOTAL FROM [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL 1] WHERE ([DIA] BETWEEN " & f1 & " AND " & f2 & ") AND ([ID_SITUACION] = 1) AND (ID_SIT_P <> 1 AND ID_SIT_P <> 23 AND ID_SIT_P <> 32) GROUP BY [D_SIT_P]"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim cmd As New SqlCommand(CADENAsql, Conexion.CxRRHH)
        cmd.CommandType = CommandType.Text
        Dim da As New SqlDataAdapter(cmd)
        Dim dt As New DataTable
        da.Fill(dt)
        Grafico1.Series(0).Points.Clear()
        Grafico1.Series(0).IsValueShownAsLabel = True
        Grafico1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Pie
        Grafico1.ChartAreas(0).Area3DStyle.Enable3D = True
        Grafico1.Series(0).Color = Color.RoyalBlue
        Grafico1.Series(0).XValueMember = "D_SIT_P"
        Grafico1.Series(0).YValueMembers = "TOTAL"
        Grafico1.DataSource = dt
        '-----------------------------------------------------------
        Me.DGV.DataSource = Nothing
        CADENAsql = "SELECT D_N1, D_N2, D_N3, D_N4, D_N5, D_N6, D_N7, D_N8, CODIGO, AN, [D_SIT_P], COUNT([ID_M_P]) AS TOTAL, ORDEN FROM [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL 1] WHERE ([DIA] BETWEEN " & f1 & " AND " & f2 & ") AND ([ID_SITUACION] = 1) AND (ID_SIT_P <> 1 AND ID_SIT_P <> 23 AND ID_SIT_P <> 32) GROUP BY D_N1, D_N2, D_N3, D_N4, D_N5, D_N6, D_N7, D_N8, CODIGO, AN, [D_SIT_P], ORDEN ORDER BY ORDEN"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE SITUACION DE PERSONAL 1]")
        DGV.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE SITUACION DE PERSONAL 1]")
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        '-----------------------------------------------------------

        Me.DGV.DefaultCellStyle.ForeColor = Color.Black

        Me.DGV.Columns(0).HeaderText = "Nivel 1"
        Me.DGV.Columns(0).Width = 40
        Me.DGV.Columns(0).Visible = True
        Me.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(1).HeaderText = "Nivel 2"
        Me.DGV.Columns(1).Width = 40
        Me.DGV.Columns(1).Visible = True
        Me.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(2).HeaderText = "Nivel 3"
        Me.DGV.Columns(2).Width = 60
        Me.DGV.Columns(2).Visible = True
        Me.DGV.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(3).HeaderText = "Nivel 4"
        Me.DGV.Columns(3).Width = 60
        Me.DGV.Columns(3).Visible = True
        Me.DGV.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(4).HeaderText = "Nivel 5"
        Me.DGV.Columns(4).Width = 60
        Me.DGV.Columns(4).Visible = True
        Me.DGV.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(5).HeaderText = "Nivel 6"
        Me.DGV.Columns(5).Width = 40
        Me.DGV.Columns(5).Visible = True
        Me.DGV.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(6).HeaderText = "Nivel 7"
        Me.DGV.Columns(6).Width = 0
        Me.DGV.Columns(6).Visible = False
        Me.DGV.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(7).HeaderText = "Nivel 8"
        Me.DGV.Columns(7).Width = 0
        Me.DGV.Columns(7).Visible = False
        Me.DGV.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(8).HeaderText = "Codigo"
        Me.DGV.Columns(8).Width = 30
        Me.DGV.Columns(8).Visible = True
        Me.DGV.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(9).HeaderText = "Colaborador"
        Me.DGV.Columns(9).Width = 60
        Me.DGV.Columns(9).Visible = True
        Me.DGV.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(10).HeaderText = "Situacion"
        Me.DGV.Columns(10).Width = 60
        Me.DGV.Columns(10).Visible = True
        Me.DGV.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(11).HeaderText = "Total"
        Me.DGV.Columns(11).Width = 30
        Me.DGV.Columns(11).Visible = True
        Me.DGV.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Me.DGV.Columns(12).HeaderText = "Orden"
        Me.DGV.Columns(12).Width = 0
        Me.DGV.Columns(12).Visible = False
    End Sub
    Private Sub CARGAR_DATA_4()
        Dim fFECHA_I As String
        fFECHA_I = Mid(Me.DateTimePicker1.Value, 1, 10)
        Dim f1 As Object = If(fFECHA_I <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_I + "', 105), 23)", "NULL")
        Dim fFECHA_I2 As String
        fFECHA_I2 = Mid(Me.DateTimePicker2.Value, 1, 10)
        Dim f2 As Object = If(fFECHA_I2 <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_I2 + "', 105), 23)", "NULL")
        CADENAsql = "SELECT [D_SIT_P], COUNT([ID_M_P]) AS TOTAL FROM [dbo].[VISTA MAESTRO DE SITUACION DE PERSONAL 1] WHERE ([DIA] BETWEEN " & f1 & " AND " & f2 & ") AND ([ID_SITUACION] = 1) GROUP BY [D_SIT_P]"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim cmd As New SqlCommand(CADENAsql, Conexion.CxRRHH)
        cmd.CommandType = CommandType.Text
        Dim da As New SqlDataAdapter(cmd)
        Dim dt As New DataTable
        da.Fill(dt)
        Grafico1.Series(0).Points.Clear()
        Grafico1.Series(0).IsValueShownAsLabel = True
        Grafico1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Pie
        Grafico1.ChartAreas(0).Area3DStyle.Enable3D = True
        Grafico1.Series(0).Color = Color.RoyalBlue
        Grafico1.Series(0).XValueMember = "D_SIT_P"
        Grafico1.Series(0).YValueMembers = "TOTAL"
        Grafico1.DataSource = dt
        '-----------------------------------------------------------
        Me.DGV.DataSource = Nothing
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE SITUACION DE PERSONAL 1]")
        DGV.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE SITUACION DE PERSONAL 1]")
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        '-----------------------------------------------------------
        Me.DGV.DefaultCellStyle.ForeColor = Color.Black

        Me.DGV.Columns(0).HeaderText = "Situacion"
        Me.DGV.Columns(0).Width = 300
        Me.DGV.Columns(0).Visible = True
        Me.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(1).HeaderText = "Total"
        Me.DGV.Columns(1).Width = 70
        Me.DGV.Columns(1).Visible = True
        Me.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
    End Sub
    Private Sub CARGAR_DATA_3()
        Dim fFECHA_I As String
        fFECHA_I = Mid(Me.DateTimePicker1.Value, 1, 10)
        Dim f1 As Object = If(fFECHA_I <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_I + "', 105), 23)", "NULL")
        Dim fFECHA_I2 As String
        fFECHA_I2 = Mid(Me.DateTimePicker2.Value, 1, 10)
        Dim f2 As Object = If(fFECHA_I2 <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_I2 + "', 105), 23)", "NULL")
        CADENAsql = "SELECT [D_SITUACION], COUNT([ID_SITUACION]) AS TOTAL  FROM [dbo].[VISTA MAESTRO DE CARGOS SIT 1] GROUP BY D_SITUACION, ID_SITUACION ORDER BY TOTAL"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim cmd As New SqlCommand(CADENAsql, Conexion.CxRRHH)
        cmd.CommandType = CommandType.Text
        Dim da As New SqlDataAdapter(cmd)
        Dim dt As New DataTable
        da.Fill(dt)
        Grafico1.Series(0).Points.Clear()
        Grafico1.Series(0).IsValueShownAsLabel = True
        Grafico1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Column
        Grafico1.ChartAreas(0).Area3DStyle.Enable3D = False
        Grafico1.Series(0).Color = Color.RoyalBlue
        Grafico1.Series(0).XValueMember = "D_SITUACION"
        Grafico1.Series(0).YValueMembers = "TOTAL"
        Grafico1.DataSource = dt
        '-----------------------------------------------------------
        Me.DGV.DataSource = Nothing
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE CARGOS SIT 1]")
        DGV.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE CARGOS SIT 1]")
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        '-----------------------------------------------------------
        Me.DGV.DefaultCellStyle.ForeColor = Color.Black

        Me.DGV.Columns(0).HeaderText = "Tipo de Situacion"
        Me.DGV.Columns(0).Width = 300
        Me.DGV.Columns(0).Visible = True
        Me.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(1).HeaderText = "Total"
        Me.DGV.Columns(1).Width = 70
        Me.DGV.Columns(1).Visible = True
        Me.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
    End Sub
    Private Sub CARGAR_DATA_2()
        Dim fFECHA_I As String
        fFECHA_I = Mid(Me.DateTimePicker1.Value, 1, 10)
        Dim f1 As Object = If(fFECHA_I <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_I + "', 105), 23)", "NULL")
        Dim fFECHA_I2 As String
        fFECHA_I2 = Mid(Me.DateTimePicker2.Value, 1, 10)
        Dim f2 As Object = If(fFECHA_I2 <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_I2 + "', 105), 23)", "NULL")
        CADENAsql = "SELECT TIPO_MARCA, COUNT (DISTINCT [ID_EMPLEADO]) AS TOTAL  FROM [dbo].[VISTA MAESTRO DE MARCAS] WHERE ([FECHA_MARCA] BETWEEN " & f1 & " AND " & f2 & ") GROUP BY TIPO_MARCA ORDER BY TOTAL DESC"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim cmd As New SqlCommand(CADENAsql, Conexion.CxRRHH)
        cmd.CommandType = CommandType.Text
        Dim da As New SqlDataAdapter(cmd)
        Dim dt As New DataTable
        da.Fill(dt)
        Grafico1.Series(0).Points.Clear()
        Grafico1.Series(0).IsValueShownAsLabel = True
        Grafico1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Pie
        Grafico1.ChartAreas(0).Area3DStyle.Enable3D = True
        Grafico1.Series(0).Color = Color.RoyalBlue
        Grafico1.Series(0).XValueMember = "TIPO_MARCA"
        Grafico1.Series(0).YValueMembers = "TOTAL"
        Grafico1.DataSource = dt
        '-----------------------------------------------------------
        Me.DGV.DataSource = Nothing
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE MARCAS]")
        DGV.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE MARCAS]")
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        '-----------------------------------------------------------
        Me.DGV.DefaultCellStyle.ForeColor = Color.Black

        Me.DGV.Columns(0).HeaderText = "Tipo de Marca"
        Me.DGV.Columns(0).Width = 300
        Me.DGV.Columns(0).Visible = True
        Me.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(1).HeaderText = "Total"
        Me.DGV.Columns(1).Width = 70
        Me.DGV.Columns(1).Visible = True
        Me.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
    End Sub
    Private Sub CARGAR_DATA_1()
        Dim fFECHA_I As String
        fFECHA_I = Mid(Me.DateTimePicker1.Value, 1, 10)
        Dim f1 As Object = If(fFECHA_I <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_I + "', 105), 23)", "NULL")
        Dim fFECHA_I2 As String
        fFECHA_I2 = Mid(Me.DateTimePicker2.Value, 1, 10)
        Dim f2 As Object = If(fFECHA_I2 <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_I2 + "', 105), 23)", "NULL")
        CADENAsql = "SELECT [RELOJ], COUNT (DISTINCT [ID_EMPLEADO]) AS TOTAL FROM [dbo].[VISTA MAESTRO DE MARCAS] WHERE ([FECHA_MARCA] BETWEEN " & f1 & " AND " & f2 & ") AND (TIPO_MARCA = 'ENTRADA') GROUP BY RELOJ ORDER BY TOTAL DESC"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim cmd As New SqlCommand(CADENAsql, Conexion.CxRRHH)
        cmd.CommandType = CommandType.Text
        Dim da As New SqlDataAdapter(cmd)
        Dim dt As New DataTable
        da.Fill(dt)
        Grafico1.Series(0).Points.Clear()
        Grafico1.Series(0).IsValueShownAsLabel = True
        Grafico1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Pie
        Grafico1.ChartAreas(0).Area3DStyle.Enable3D = True
        Grafico1.Series(0).Color = Color.RoyalBlue
        Grafico1.Series(0).XValueMember = "RELOJ"
        Grafico1.Series(0).YValueMembers = "TOTAL"
        Grafico1.DataSource = dt
        '-----------------------------------------------------------
        Me.DGV.DataSource = Nothing
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE MARCAS]")
        DGV.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE MARCAS]")
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        '-----------------------------------------------------------
        Me.DGV.DefaultCellStyle.ForeColor = Color.Black

        Me.DGV.Columns(0).HeaderText = "Reloj"
        Me.DGV.Columns(0).Width = 300
        Me.DGV.Columns(0).Visible = True
        Me.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(1).HeaderText = "Total"
        Me.DGV.Columns(1).Width = 70
        Me.DGV.Columns(1).Visible = True
        Me.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Call SELECCION()
        Call CONTAR()
    End Sub
    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Call SELECCION()
        Call CONTAR()
    End Sub
    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        Call SELECCION()
        Call CONTAR()
    End Sub
    Private Sub SELECCION()
        Application.DoEvents()
        Me.Cursor = Cursors.WaitCursor
        'Asistencia Total del Personal por Rango de Fecha y Reloj
        If Me.ComboBox1.SelectedValue = 1 Then
            Call CARGAR_DATA_1()
        End If
        'Asistencia Total del Personal por Rango de Fecha y Tipo (Entradas\ Salidas)
        If Me.ComboBox1.SelectedValue = 2 Then
            Call CARGAR_DATA_2()
        End If
        'Total de Cargos por Situacion de Cargos (Ocupados y Vacantes)
        If Me.ComboBox1.SelectedValue = 3 Then
            Call CARGAR_DATA_3()
        End If
        'Total de Personal por Tipo de Situacion
        If Me.ComboBox1.SelectedValue = 4 Then
            Call CARGAR_DATA_4()
        End If
        'Total de Dias del Personal en Situacion por Niveles Organizativos y Tipo de Situacion
        If Me.ComboBox1.SelectedValue = 5 Then
            Call CARGAR_DATA_5()
        End If
        'Total de Dias del Personal en Situacion por Niveles Organizativos y Tipo de Situacion (Nivel 3)
        If Me.ComboBox1.SelectedValue = 6 Then
            Call CARGAR_DATA_6()
        End If
        'Total de Dias del Personal en Situacion por Niveles Organizativos y Tipo de Situacion (Nivel 4)
        If Me.ComboBox1.SelectedValue = 7 Then
            Call CARGAR_DATA_7()
        End If
        'Total de Dias del Personal en Situacion por Niveles Organizativos y Tipo de Situacion (Nivel 5)
        If Me.ComboBox1.SelectedValue = 8 Then
            Call CARGAR_DATA_8()
        End If
        Me.Cursor = Cursors.Default
    End Sub
    Dim C As Integer
    Private Sub CONTAR()
        'Asistencia Total del Personal por Rango de Fecha y Reloj
        If Me.ComboBox1.SelectedValue = 1 Then
            C = 1
            GoTo RECORRER
        End If
        'Asistencia Total del Personal por Rango de Fecha y Tipo (Entradas\ Salidas)
        If Me.ComboBox1.SelectedValue = 2 Then
            C = 1
            GoTo RECORRER
        End If
        'Total de Cargos por Situacion de Cargos (Ocupados y Vacantes)
        If Me.ComboBox1.SelectedValue = 3 Then
            C = 1
            GoTo RECORRER
        End If
        'Total de Personal por Tipo de Situacion
        If Me.ComboBox1.SelectedValue = 4 Then
            C = 1
            GoTo RECORRER
        End If
        'Total de Dias del Personal en Situacion por Niveles Organizativos y Tipo de Situacion
        If Me.ComboBox1.SelectedValue = 5 Then
            C = 11
            GoTo RECORRER
        End If
        'Total de Dias del Personal en Situacion por Niveles Organizativos y Tipo de Situacion (Nivel 3)
        If Me.ComboBox1.SelectedValue = 6 Then
            C = 4
            GoTo RECORRER
        End If
        'Total de Dias del Personal en Situacion por Niveles Organizativos y Tipo de Situacion (Nivel 4)
        If Me.ComboBox1.SelectedValue = 7 Then
            C = 5
            GoTo RECORRER
        End If
        'Total de Dias del Personal en Situacion por Niveles Organizativos y Tipo de Situacion (Nivel 5)
        If Me.ComboBox1.SelectedValue = 8 Then
            C = 6
            GoTo RECORRER
        End If
RECORRER:
        On Error GoTo SALIR
        Dim CANTIDAD As Integer
        CANTIDAD = 0
        If Me.DGV.RowCount > 0 Then
            For I = 0 To Me.DGV.RowCount - 1
                DGV.Rows(I).Selected = True
                Application.DoEvents()
                DGV.CurrentCell = Me.DGV.Item(C, I)
                CANTIDAD = CANTIDAD + Me.DGV.Item(C, I).Value
            Next
        End If
        Me.Label2.Text = "Total ---> " & CANTIDAD
        Exit Sub
SALIR:
    End Sub
    Private Sub Frm_Estadisticas_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        ESTADISTICAS = False
    End Sub
    Private Sub Frm_Estadisticas_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        ESTADISTICAS = False
    End Sub
    Private Sub Frm_Estadisticas_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        ESTADISTICAS = False
    End Sub
    Private Sub Frm_Estadisticas_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        ESTADISTICAS = False
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            TITULO_EXCEL = Me.ComboBox1.Text
            EXPORTAR_DATOS_A_EXCEL(DGV, "Desde: " & Mid(Me.DateTimePicker1.Value, 1, 10) & " Hasta: " & Mid(Me.DateTimePicker2.Value, 1, 10) & " - " & Me.Label2.Text)
        Catch ex As Exception
            MessageBox.Show("No se puede generar el documento Excel.", "Mensaje del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class