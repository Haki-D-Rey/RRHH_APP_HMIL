Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Public Class Frm_Consulta_General
    Private Sub Frm_Consulta_General_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            CONSULTA_GENERAL = False
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Consulta_General_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.CheckBox1.Checked = True
        Me.RadioButton2.Checked = True
        Me.CheckedListBox1.Items.Clear()
        CADENAsql = "Select ROW_NUMBER() OVER (ORDER BY COL.colid) As id_columna, " &
                    "COL.name AS columna " &
                    "From dbo.syscolumns COL " &
                    "Join dbo.sysobjects OBJ On OBJ.id = COL.id " &
                    "Join dbo.systypes TYP ON TYP.xusertype = COL.xtype " &
                    "Left Join dbo.sysforeignkeys FK ON FK.fkey = COL.colid And FK.fkeyid=OBJ.id " &
                    "Left Join dbo.sysobjects OBJ2 ON OBJ2.id = FK.rkeyid " &
                    "Left Join dbo.syscolumns COL2 ON COL2.colid = FK.rkey And COL2.id = OBJ2.id " &
                    "WHERE OBJ.name = 'VISTA CONSULTA GENERAL DE EXPEDIENTES FINAL' AND (OBJ.xtype='U' OR OBJ.xtype='V')"
        Call CARGAR_LISTA_COLUMNAS()
        Me.DGV1.DataSource = Nothing
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        CONSULTA_GENERAL = False
        Me.Close()
    End Sub
    Private Sub CARGAR_LISTA_COLUMNAS()
        Dim I As Integer
        Me.Cursor = Cursors.WaitCursor
        I = 0
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                Application.DoEvents()
                For I = 0 To MiDataTable.Rows.Count - 1
                    Me.CheckedListBox1.Items.Add(MiDataTable.Rows(I).Item(1).ToString)
                Next
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        For i As Integer = 0 To CheckedListBox1.Items.Count - 1
            CheckedListBox1.SetItemChecked(i, True)
        Next
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        For i As Integer = 0 To CheckedListBox1.Items.Count - 1
            CheckedListBox1.SetItemChecked(i, False)
        Next
    End Sub
    Private Sub ARMAR_CADENA_CAMPOS()
        CAMPOS = ""
        If CheckedListBox1.CheckedItems.Count <> 0 Then
            For i As Integer = 0 To CheckedListBox1.Items.Count - 1
                If CheckedListBox1.GetItemChecked(i) Then
                    If Me.CheckedListBox1.Items(i).ToString = "ID" Or Me.CheckedListBox1.Items(i).ToString = "ID_ESTADO_P" Then
                        GoTo SALTO
                    End If
                    If Me.CheckBox1.Checked = True Then
                        If CAMPOS = "" Then
                            CAMPOS = "[" & Me.CheckedListBox1.Items(i).ToString & "]"
                        Else
                            CAMPOS = CAMPOS & ", [" & Me.CheckedListBox1.Items(i).ToString & "]"
                        End If
                    Else
                        If CAMPOS = "" Then
                            CAMPOS = "dbo.ProperCase([" & Me.CheckedListBox1.Items(i).ToString & "]) AS [" & Me.CheckedListBox1.Items(i).ToString & "]"
                        Else
                            CAMPOS = CAMPOS & ", dbo.ProperCase([" & Me.CheckedListBox1.Items(i).ToString & "]) AS [" & Me.CheckedListBox1.Items(i).ToString & "]"
                        End If
                    End If
SALTO:
                End If
            Next
        End If
        If CAMPOS = "" Then
            CAMPOS = "*"
        End If
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW()
        Application.DoEvents()
        Me.Cursor = Cursors.WaitCursor
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql1, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA CONSULTA GENERAL DE EXPEDIENTES FINAL]")
        DGV1.DataSource = MiDataSet.Tables("[VISTA CONSULTA GENERAL DE EXPEDIENTES FINAL]")
        Me.Label4.Text = DGV1.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Me.Cursor = Cursors.Default
    End Sub
    Dim CAMPOS As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call VERIFICAR_ACCESOS("301") : If HAY_ACCESO = False Then : Exit Sub : End If
        If Me.RadioButton1.Checked = False And Me.RadioButton2.Checked = False Then
            MsgBox("Debe seleccionar el Tipo de Estructura primeramente", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Call ARMAR_CADENA_CAMPOS()
        If CAMPOS = "" Then
            MsgBox("Debe seleccionar las columnas que desea visualizar primeramente", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        SELECCION = ""
        If Me.RadioButton1.Checked = True Then
            If SELECCION = "" Then
                SELECCION = "[CARGO MILITAR] = 1"
            Else
                SELECCION = SELECCION & " AND [CARGO MILITAR] = 1'"
            End If
        End If
        If Me.RadioButton2.Checked = True Then
            If SELECCION = "" Then
                SELECCION = "[CARGO PAME] = 1"
            Else
                SELECCION = SELECCION & " AND [CARGO PAME] = 1'"
            End If
        End If
        If SELECCION = "" Then
            SELECCION = "[SITUACION DEL CARGO]= 'OCUPADO'"
        Else
            SELECCION = SELECCION & " AND [SITUACION DEL CARGO] = 'OCUPADO'"
        End If
        CADENAsql1 = "SELECT " & CAMPOS & " FROM [VISTA CONSULTA GENERAL DE EXPEDIENTES FINAL] WHERE " & SELECCION & ""
        Call CARGAR_DATAGRIDVIEW()
        For I = 0 To Me.DGV1.ColumnCount - 1
            Me.DGV1.Columns(I).Width = 70
            Me.DGV1.Columns(I).Visible = True
            Me.DGV1.Columns(I).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            Me.DGV1.Columns(I).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Next
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call VERIFICAR_ACCESOS("302") : If HAY_ACCESO = False Then : Exit Sub : End If
        If Me.RadioButton1.Checked = False And Me.RadioButton2.Checked = False Then
            MsgBox("Debe seleccionar el Tipo de Estructura primeramente", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Try
            TITULO_EXCEL = "CONSULTA GENERAL"
            If Me.RadioButton1.Checked = True Then
                EXPORTAR_DATOS_A_EXCEL(DGV1, "Consulta General [Estructura Militar] (" & UCase(Now) & ")")
            End If
            If Me.RadioButton2.Checked = True Then
                EXPORTAR_DATOS_A_EXCEL(DGV1, "Consulta General [Estructura Pame] (" & UCase(Now) & ")")
            End If
        Catch ex As Exception
            MessageBox.Show("No se puede generar el documento Excel.", "Mensaje del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class