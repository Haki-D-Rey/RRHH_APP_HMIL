Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Expedientes_Inactivos_Agregar
    Private Sub Frm_Expedientes_Inactivos_Agregar_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = ""
            Me.TextBox2.Text = "0"
            Me.ComboBox1.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Expedientes_Inactivos_Agregar_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = ""
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_DOCUMENTO()
        Me.TextBox2.Text = ""
        Me.ComboBox1.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = "0"
        Me.ComboBox1.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_DOCUMENTO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE DOCUMENTO] order by ID_T_DOC", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_T_DOC"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
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
    Dim ARCHIVO_DE_ORIGEN As String
    Dim ARCHIVO_DE_DESTINO As String
    Dim UBICACION_DE_ORIGEN As String
    Dim UBICACION_DE_DESTINO As String
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.OpenFileDialog.Filter = "Archivos PDF (*.pdf)|*.pdf"
        Me.OpenFileDialog.FileName = ""
        Me.OpenFileDialog.ShowDialog()
        ARCHIVO_DE_ORIGEN = UCase(Trim(Me.OpenFileDialog.SafeFileName))
        UBICACION_DE_ORIGEN = UCase(Trim(Me.OpenFileDialog.FileName))
        If ARCHIVO_DE_ORIGEN <> "" Then
            Me.TextBox2.Text = ARCHIVO_DE_ORIGEN
        End If
        Me.TextBox2.Focus()
    End Sub
    Private Sub GENERAR_CODIGO()
        CODIGO = 0
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            MiDataAdapter = New SqlDataAdapter("SELECT ID_M_DD FROM [MAESTRO DE DOCUMENTOS DIGITALES] ORDER BY ID_M_DD", Conexion.CxRRHH)
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
            Me.TextBox1.Text = CODIGO
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim vcMIXTA, vcMILITAR, vcPAME As Boolean
    Private Sub BUSCAR_ESTRUCTURA_CARGO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [MAESTRO DE CARGOS] WHERE ID_M_C = '" & Frm_Expedientes_Inactivos.TextBox36.Text & "'", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                vcMIXTA = MiDataTable.Rows(0).Item(15).ToString
                vcMILITAR = MiDataTable.Rows(0).Item(16).ToString
                vcPAME = MiDataTable.Rows(0).Item(17).ToString
            Else
                vcMIXTA = False
                vcMILITAR = False
                vcPAME = False
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.TextBox1.Text = "0"
        Call GENERAR_CODIGO()
        If TextBox1.Text = "0" Or TextBox1.Text = "" Then
            MsgBox("No se ha generado el Codigo del registro", vbInformation, "Mensaje del Sistema")
            Me.TextBox1.Text = "0"
            Me.Button2.Focus()
            Exit Sub
        End If
        If ARCHIVO_DE_ORIGEN = "" Or UBICACION_DE_ORIGEN = "" Then
            MsgBox("Debe seleccionar el Archivo Digital Correctamente", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Tipo de Documento Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox2.Text) = 0 Then
            Me.TextBox2.Text = "."
        End If
        Call BUSCAR_DIRECTORIO()
        If idDIRECTORIO <> 0 Then
            Call COPIAR_ARCHIVO_SERVIDOR()
            ARCHIVO_DE_DESTINO = Frm_Expedientes_Inactivos.TextBox2.Text & "-" & Me.TextBox1.Text & "-" & Me.ComboBox1.Text & ".pdf"
            Call RENOMBRAR_ARCHIVO_SERVIDOR()
            Call GRABAR_REGISTRO()
            vmdID_M_DD = Me.TextBox1.Text
            CADENAsql = "Select * FROM [VISTA MAESTRO DE DOCUMENTOS DIGITALES] WHERE ID_M_P = " & Frm_Expedientes_Inactivos.TextBox35.Text & " ORDER BY ID_M_P, ID_T_DOC"
            Call CARGAR_DATAGRIDVIEW_12()
            Call ARMAR_DATAGRIDVIEW_12()
            Call SELECCIONAR_REGISTRO_DATAGRIDVIEW_12()
        Else
            MsgBox("No se encuentra el Directorio de Documentos Digitales", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
    End Sub
    Private Sub COPIAR_ARCHIVO_SERVIDOR()
        My.Computer.FileSystem.CopyFile(UBICACION_DE_ORIGEN, dDIRECTORIO & "\" & ARCHIVO_DE_ORIGEN)
    End Sub
    Private Sub RENOMBRAR_ARCHIVO_SERVIDOR()
        My.Computer.FileSystem.RenameFile(dDIRECTORIO & "\" & ARCHIVO_DE_ORIGEN, ARCHIVO_DE_DESTINO)
    End Sub
    Dim idDIRECTORIO As Integer
    Dim dDIRECTORIO As String
    Private Sub BUSCAR_DIRECTORIO()
        Dim xSqlDataAdapter As SqlDataAdapter
        Dim xDataTable As New DataTable
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Try
            Conexion.ABD_RRHH()
            xSqlDataAdapter = New SqlDataAdapter("Select * FROM [DIRECTORIO DE DOCUMENTOS DIGITALES] WHERE ID_D1 = '1'", Conexion.CxRRHH)
            xDataTable.Clear()
            xSqlDataAdapter.Fill(xDataTable)
            If xDataTable.Rows.Count > 0 Then
                idDIRECTORIO = xDataTable.Rows(0).Item(0).ToString
                dDIRECTORIO = xDataTable.Rows(0).Item(1).ToString
            Else
                idDIRECTORIO = 0
                dDIRECTORIO = ""
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub GRABAR_REGISTRO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE DOCUMENTOS DIGITALES] (ID_M_DD, ID_M_P, ID_D1, ID_T_DOC, NOMBRE, OBSERVACION)" &
                                  "values (" & CInt(Me.TextBox1.Text) & ", " & Frm_Expedientes_Inactivos.TextBox35.Text & ", 1, " & Me.ComboBox1.SelectedValue & ", '" & ARCHIVO_DE_DESTINO & "', '" & Me.TextBox2.Text & "')"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_12()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE DOCUMENTOS DIGITALES]")
        Frm_Expedientes_Inactivos.DGV12.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE DOCUMENTOS DIGITALES]")
        Frm_Expedientes_Inactivos.Label90.Text = "Total de Registros: " & Frm_Expedientes_Inactivos.DGV12.RowCount
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_12()
        Frm_Expedientes_Inactivos.DGV12.Columns(0).HeaderText = "Id"
        Frm_Expedientes_Inactivos.DGV12.Columns(0).Width = 30
        Frm_Expedientes_Inactivos.DGV12.Columns(0).Visible = True
        Frm_Expedientes_Inactivos.DGV12.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Expedientes_Inactivos.DGV12.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Frm_Expedientes_Inactivos.DGV12.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Frm_Expedientes_Inactivos.DGV12.Columns(1).HeaderText = "ID_M_P"
        Frm_Expedientes_Inactivos.DGV12.Columns(1).Width = 0
        Frm_Expedientes_Inactivos.DGV12.Columns(1).Visible = False

        Frm_Expedientes_Inactivos.DGV12.Columns(2).HeaderText = "ID_D1"
        Frm_Expedientes_Inactivos.DGV12.Columns(2).Width = 0
        Frm_Expedientes_Inactivos.DGV12.Columns(2).Visible = False

        Frm_Expedientes_Inactivos.DGV12.Columns(3).HeaderText = "DIRECTORIO"
        Frm_Expedientes_Inactivos.DGV12.Columns(3).Width = 0
        Frm_Expedientes_Inactivos.DGV12.Columns(3).Visible = False

        Frm_Expedientes_Inactivos.DGV12.Columns(4).HeaderText = "ID_T_DOC"
        Frm_Expedientes_Inactivos.DGV12.Columns(4).Width = 0
        Frm_Expedientes_Inactivos.DGV12.Columns(4).Visible = False

        Frm_Expedientes_Inactivos.DGV12.Columns(5).HeaderText = "Tipo de Documento"
        Frm_Expedientes_Inactivos.DGV12.Columns(5).Width = 200
        Frm_Expedientes_Inactivos.DGV12.Columns(5).Visible = True
        Frm_Expedientes_Inactivos.DGV12.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Inactivos.DGV12.Columns(6).HeaderText = "Nombre"
        Frm_Expedientes_Inactivos.DGV12.Columns(6).Width = 300
        Frm_Expedientes_Inactivos.DGV12.Columns(6).Visible = True
        Frm_Expedientes_Inactivos.DGV12.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Expedientes_Inactivos.DGV12.Columns(7).HeaderText = "Observacion"
        Frm_Expedientes_Inactivos.DGV12.Columns(7).Width = 290
        Frm_Expedientes_Inactivos.DGV12.Columns(7).Visible = True
        Frm_Expedientes_Inactivos.DGV12.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW_12()
        Dim I As Integer
        For I = 0 To Frm_Expedientes_Inactivos.DGV12.RowCount - 1
            If Frm_Expedientes_Inactivos.DGV12.Item(0, I).Value = vmdID_M_DD Then
                Frm_Expedientes_Inactivos.DGV12.Rows(I).Selected = True
                Frm_Expedientes_Inactivos.DGV12.CurrentCell = Frm_Expedientes_Inactivos.DGV12.Item(0, I)
                Exit Sub
            End If
        Next
    End Sub
End Class