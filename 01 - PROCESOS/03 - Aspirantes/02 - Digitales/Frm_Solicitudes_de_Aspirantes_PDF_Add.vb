Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Solicitudes_de_Aspirantes_PDF_Add
    Private Sub Frm_Solicitudes_de_Aspirantes_PDF_Add_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = ""
            Me.TextBox2.Text = "0"
            Me.ComboBox1.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Solicitudes_de_Aspirantes_PDF_Add_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = ""
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_DOCUMENTO_ASPIRANTES()
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
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_DOCUMENTO_ASPIRANTES()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE DOCUMENTO ASPIRANTES] order by ID_T_DOC_ASP", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_T_DOC_ASP"
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
            MiDataAdapter = New SqlDataAdapter("SELECT ID_M_DD_A FROM [MAESTRO DE DOCUMENTOS DIGITALES ASPIRANTES] ORDER BY ID_M_DD_A", Conexion.CxRRHH)
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
            ARCHIVO_DE_DESTINO = "ASPIRANTE_" & Format(Now.Day, "00") & Format(Now.Month, "00") & Format(Now.Year, "0000") & "_" & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00") & ".pdf"
            Call RENOMBRAR_ARCHIVO_SERVIDOR()
            Call COPIAR_ARCHIVO_SERVIDOR()
            Call GRABAR_REGISTRO()
            vmdID_M_DD = Me.TextBox1.Text
            CERRAR = False
            Me.Close()
        Else
            MsgBox("No se encuentra el Directorio de Documentos Digitales", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
    End Sub
    Private Sub COPIAR_ARCHIVO_SERVIDOR()
        Dim NUEVA_RUTA As String
        NUEVA_RUTA = Mid(UBICACION_DE_ORIGEN, 1, Len(UBICACION_DE_ORIGEN) - Len(ARCHIVO_DE_ORIGEN)) & ARCHIVO_DE_DESTINO
        My.Computer.FileSystem.CopyFile(NUEVA_RUTA, dDIRECTORIO & "\" & ARCHIVO_DE_DESTINO)
    End Sub
    Private Sub RENOMBRAR_ARCHIVO_SERVIDOR()
        My.Computer.FileSystem.RenameFile(UBICACION_DE_ORIGEN, ARCHIVO_DE_DESTINO)
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
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE DOCUMENTOS DIGITALES ASPIRANTES] (ID_M_DD_A, ID_S, ID_D1, ID_T_DOC_ASP, NOMBRE, OBSERVACION)" &
                                  "values (" & CInt(Me.TextBox1.Text) & ", " & aspID_S & ", 1, " & Me.ComboBox1.SelectedValue & ", '" & ARCHIVO_DE_DESTINO & "', '" & Me.TextBox2.Text & "')"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
End Class