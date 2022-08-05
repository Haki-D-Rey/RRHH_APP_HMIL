Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Verificar_Visualizar
    Private Sub Frm_Verificar_Visualizar_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Verificar_Visualizar_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If vTIPO = "COBERTURA MILITAR" Then
            Me.Text = "Detalles: Identidad [" & vIDENTIDAD & "], Nombres y Apellidos [" & vNOMBRES_Y_APELLIDOS & "], Tipo de Cobertura [" & vTIPO & "]"
            CADENAsql = "Select * from [SIE VISTA MAESTRO DE CONSULTA PRINCIPAL] WHERE (IDENTIDAD = '" & vIDENTIDAD & "') ORDER BY [TIPO PARENTELA] DESC"
        End If
        If vTIPO = "ASEGURADO" Or vTIPO = "RIESGO PROFESIONAL" Or vTIPO = "PENSIONADOS" Then
            Me.Text = "Detalles: Nss [" & vNSS & "], Nombres y Apellidos [" & vNOMBRES_Y_APELLIDOS & "], Tipo de Cobertura [" & vTIPO & "]"
            CADENAsql = "Select * from [SIE VISTA MAESTRO DE CONSULTA PRINCIPAL] WHERE (NSS = '" & vNSS & "') ORDER BY [TIPO] DESC"
        End If
        'If vTIPO = "PROGRAMAS ESPECIALES" Then
        '    Me.Text = "Detalles: Nss [" & vNSS & "], Nombres y Apellidos [" & vNOMBRES_Y_APELLIDOS & "], Tipo de Cobertura [" & vTIPO & "]"
        '    CADENAsql = "Select * from [SIE VISTA MAESTRO DE CONSULTA PRINCIPAL] WHERE (NO_CARNE = '" & vNO_CARNE & "') ORDER BY [TIPO] DESC"
        'End If
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
        Me.DGV.Focus()
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW()
        If Conexion.CxSIE.State = ConnectionState.Open Then
            Conexion.CBD_SIE()
        End If
        Conexion.ABD_SIE()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxSIE)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[SIE VISTA MAESTRO DE CONSULTA PRINCIPAL]")
        DGV.DataSource = MiDataSet.Tables("[SIE VISTA MAESTRO DE CONSULTA PRINCIPAL]")
        If Conexion.CxSIE.State = ConnectionState.Open Then
            Conexion.CBD_SIE()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW()
        Me.DGV.Columns(0).HeaderText = "Tipo de Cobertura"
        Me.DGV.Columns(0).Width = 150
        Me.DGV.Columns(0).Visible = True
        Me.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(1).HeaderText = "Tipo de Ingreso"
        Me.DGV.Columns(1).Width = 100
        Me.DGV.Columns(1).Visible = True
        Me.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(2).HeaderText = "Nss"
        Me.DGV.Columns(2).Width = 60
        Me.DGV.Columns(2).Visible = True
        Me.DGV.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(3).HeaderText = "Nombres"
        Me.DGV.Columns(3).Width = 0
        Me.DGV.Columns(3).Visible = False
        Me.DGV.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(4).HeaderText = "Apellidos"
        Me.DGV.Columns(4).Width = 0
        Me.DGV.Columns(4).Visible = False
        Me.DGV.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(5).HeaderText = "Nombres y Apellidos"
        Me.DGV.Columns(5).Width = 300
        Me.DGV.Columns(5).Visible = True
        Me.DGV.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(6).HeaderText = ""
        Me.DGV.Columns(6).Width = 10
        Me.DGV.Columns(6).Visible = True
        Me.DGV.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(6).DefaultCellStyle.BackColor = Color.Blue
        Me.DGV.Columns(6).Frozen = True

        Me.DGV.Columns(7).HeaderText = "Identidad (8 Digitos)"
        Me.DGV.Columns(7).Width = 80
        Me.DGV.Columns(7).Visible = True
        Me.DGV.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(8).HeaderText = "Sexo"
        Me.DGV.Columns(8).Width = 40
        Me.DGV.Columns(8).Visible = True
        Me.DGV.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(9).HeaderText = "Grado Militar"
        Me.DGV.Columns(9).Width = 150
        Me.DGV.Columns(9).Visible = True
        Me.DGV.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(10).HeaderText = "Tipo de Parentela"
        Me.DGV.Columns(10).Width = 80
        Me.DGV.Columns(10).Visible = True
        Me.DGV.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(11).HeaderText = "Unidad Militar"
        Me.DGV.Columns(11).Width = 60
        Me.DGV.Columns(11).Visible = True
        Me.DGV.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(12).HeaderText = "Cargo Militar"
        Me.DGV.Columns(12).Width = 150
        Me.DGV.Columns(12).Visible = True
        Me.DGV.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(13).HeaderText = "Identidad (13 Digitos)"
        Me.DGV.Columns(13).Width = 100
        Me.DGV.Columns(13).Visible = True
        Me.DGV.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(14).HeaderText = "Clasificacion"
        Me.DGV.Columns(14).Width = 80
        Me.DGV.Columns(14).Visible = True
        Me.DGV.Columns(14).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(14).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(15).HeaderText = "No. Cedula"
        Me.DGV.Columns(15).Width = 120
        Me.DGV.Columns(15).Visible = True
        Me.DGV.Columns(15).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(15).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(16).HeaderText = "No. Carne Beneficiario"
        Me.DGV.Columns(16).Width = 80
        Me.DGV.Columns(16).Visible = True
        Me.DGV.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(16).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(17).HeaderText = "Fecha de Ingreso"
        Me.DGV.Columns(17).Width = 75
        Me.DGV.Columns(17).Visible = True
        Me.DGV.Columns(17).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(17).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(18).HeaderText = "Edad"
        Me.DGV.Columns(18).Width = 40
        Me.DGV.Columns(18).Visible = True
        Me.DGV.Columns(18).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(18).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(19).HeaderText = "Fecha de Nacimiento"
        Me.DGV.Columns(19).Width = 75
        Me.DGV.Columns(19).Visible = True
        Me.DGV.Columns(19).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(19).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub DGV_RowPrePaint(sender As Object, e As DataGridViewRowPrePaintEventArgs) Handles DGV.RowPrePaint
        Call FORMATO_FILAS_DATAGRIDVIEW()
    End Sub
    Private Sub FORMATO_FILAS_DATAGRIDVIEW()
        Dim style As New DataGridViewCellStyle
        style.Font = New Font(DGV.Font, FontStyle.Bold)
        For Each Row As DataGridViewRow In DGV.Rows
            If Row.Cells(0).Value = "COBERTURA MILITAR" Then
                Row.DefaultCellStyle.ForeColor = Color.DarkGreen
            End If
            If Row.Cells(0).Value = "ASEGURADO" Then
                Row.DefaultCellStyle.ForeColor = Color.Blue
            End If
            If Row.Cells(0).Value = "RIESGO PROFESIONAL" Then
                Row.DefaultCellStyle.ForeColor = Color.Brown
            End If
            If Row.Cells(0).Value = "PENSIONADOS" Then
                Row.DefaultCellStyle.BackColor = Color.DarkRed
            End If
        Next
    End Sub
End Class