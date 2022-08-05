Imports Microsoft.VisualBasic
Imports System
Imports System.Windows.Forms
Imports System.IO
Imports System.Diagnostics
Imports System.Data.SqlClient
Imports System.Data.OleDb
Public Class Frm_Act_Documentos_Digitales
    Private cancelar As Boolean = False
    Private Sub Frm_Act_Documentos_Digitales_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If Me.WindowState = FormWindowState.Normal Then
            My.Settings.Location = Me.Location
            My.Settings.Size = Me.Size
        Else
            My.Settings.Location = Me.RestoreBounds.Location
            My.Settings.Size = Me.RestoreBounds.Size
        End If
        My.Settings.Save()
    End Sub
    Private Sub Frm_Act_Documentos_Digitales_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        With My.Settings
            If .Location.X > -1 Then
                Me.Location = .Location
                Me.Size = .Size
            End If
            Me.txtDir.Text = .Directorio
            Me.txtFiltro.Text = .Filtro
            Me.chkConSubDir.Checked = .conSubDir
            Me.chkIgnorarError.Checked = .IgnorarError
        End With
        With My.Application.Info

        End With

        Me.btnAbrirDir.Enabled = False
        Me.btnAbrirFic.Enabled = False
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_DOCUMENTO
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_DOCUMENTO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE DOCUMENTO] ORDER BY ID_T_DOC", Conexion.CxRRHH)
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
    Private Sub recorrerDir(ByVal di As DirectoryInfo)
        Dim fics() As FileInfo
        Dim dirs() As DirectoryInfo
        Application.DoEvents()
        If cancelar Then Exit Sub
        Me.LabelInfo.Text = di.FullName & "..."
        Me.LabelInfo.Refresh()
        With My.Settings
            Try
                fics = di.GetFiles(.Filtro, SearchOption.TopDirectoryOnly)
                With Me.lvFics
                    For Each fi As FileInfo In fics
                        Dim lvi As ListViewItem = .Items.Add(fi.Name)
                        lvi.SubItems.Add(fi.DirectoryName)
                    Next
                End With
                If .conSubDir Then
                    dirs = di.GetDirectories()
                    For Each dir As DirectoryInfo In dirs
                        recorrerDir(dir)
                    Next
                End If
            Catch ex As Exception
                If .IgnorarError Then
                    Exit Sub
                End If

                If MessageBox.Show("Error: " & ex.Message & vbCrLf &
                                   "¿Quieres cancelar o continuar?",
                                   "Buscar en directorios",
                                   MessageBoxButtons.OKCancel,
                                   MessageBoxIcon.Exclamation
                                   ) = DialogResult.Cancel Then
                    cancelar = True
                    Application.DoEvents()
                End If
            End Try
        End With
    End Sub
    '' Para usar si asignamos True a HotTracking (y HoverSelection)
    '' que se cambia el elemento a un link cuando se posiciona
    '' el ratón encima.
    'Private Sub lvFics_MouseClick(ByVal sender As Object, _
    '                              ByVal e As MouseEventArgs) _
    '                              Handles lvFics.MouseClick
    '    With lvFics
    '        If .SelectedIndices.Count = 0 Then
    '            Exit Sub
    '        End If
    '        ' Comprobar en que columna se ha hecho doble clic
    '        ' El valor de e.X es relativo al control,
    '        ' por tanto, no hace falta añadir nada más.
    '        If e.X < .Columns(0).Width Then
    '            ' El nombre
    '            ' Abrir el fichero indicado
    '            ' Combinar los paths para que se agregue el separador de directorio
    '            ' si así hiciera falta
    '            Dim fic As String = Path.Combine(.SelectedItems(0).SubItems(1).Text, _
    '                                             .SelectedItems(0).SubItems(0).Text)
    '            Process.Start(fic)
    '        Else
    '            ' El directorio
    '            ' Abrir la ventana con el directorio del fichero indicado
    '            Dim dir As String = .SelectedItems(0).SubItems(1).Text
    '            Process.Start("explorer.exe", dir)
    '        End If
    '    End With
    'End Sub
    'Private Sub lvFics_DoubleClick(ByVal sender As Object, _
    '                               ByVal e As EventArgs) _
    '                               Handles lvFics.DoubleClick
    '    With lvFics
    '        If .SelectedIndices.Count > 0 Then
    '            ' Abrir la ventana con el directorio del fichero indicado
    '            Dim dir As String = .SelectedItems(0).SubItems(1).Text
    '            Process.Start("explorer.exe", dir)
    '        End If
    '    End With
    'End Sub
    Private Sub lvFics_MouseDoubleClick(ByVal sender As Object,
                                        ByVal e As MouseEventArgs) _
                                        Handles lvFics.MouseDoubleClick
        With lvFics
            If .SelectedIndices.Count = 0 Then
                Exit Sub
            End If
            If e.X < .Columns(0).Width Then
                Dim fic As String = Path.Combine(.SelectedItems(0).SubItems(1).Text,
                                                 .SelectedItems(0).SubItems(0).Text)
                Process.Start(fic)
            Else
                Dim dir As String = .SelectedItems(0).SubItems(1).Text
                Process.Start("explorer.exe", dir)
            End If

        End With
    End Sub
    Private Sub btnExaminar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnExaminar.Click
        Dim oFD As New OpenFileDialog
        With oFD
            .ValidateNames = False
            .CheckFileExists = False
            .CheckPathExists = True
            .FileName = "Folder Selection."
            If .ShowDialog = DialogResult.OK Then
                Me.txtDir.Text = Path.GetDirectoryName(.FileName)
            End If
        End With
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Static yaEstoy As Boolean
        If yaEstoy Then
            cancelar = True
            Me.btnBuscar.Text = "Cancelando..."
            Application.DoEvents()
            Exit Sub
        End If
        yaEstoy = True
        With My.Settings
            .Location = Me.Location
            .Size = Me.Size
            .Directorio = Me.txtDir.Text
            .Filtro = Me.txtFiltro.Text
            .conSubDir = Me.chkConSubDir.Checked
            .IgnorarError = Me.chkIgnorarError.Checked
            .Save()
            Dim di As New DirectoryInfo(.Directorio)
            Me.lvFics.Items.Clear()
            Me.LabelInfo.Text = "Buscando los ficheros..."
            Me.Cursor = Cursors.AppStarting
            Me.btnBuscar.Text = "Buscando..."
            Me.Refresh()
            recorrerDir(di)
        End With
        Me.Cursor = Cursors.Default
        Me.LabelInfo.Text = "Se han hallado " & Me.lvFics.Items.Count & " ficheros"
        Me.btnBuscar.Text = "Buscar"
        If Me.lvFics.Items.Count > 0 Then
            Me.btnAbrirDir.Enabled = True
            Me.btnAbrirFic.Enabled = True
        End If
        Me.Refresh()
        cancelar = False
        yaEstoy = False
    End Sub

    Private Sub btnAbrirFic_Click(sender As Object, e As EventArgs) Handles btnAbrirFic.Click
        With lvFics
            If .SelectedIndices.Count = 0 Then
                Exit Sub
            End If
            Dim fic As String = Path.Combine(.SelectedItems(0).SubItems(1).Text,
                                             .SelectedItems(0).SubItems(0).Text)
            Process.Start(fic)
        End With
    End Sub
    Private Sub btnAbrirDir_Click(sender As Object, e As EventArgs) Handles btnAbrirDir.Click
        With lvFics
            If .SelectedIndices.Count = 0 Then
                Exit Sub
            End If
            Dim dir As String = .SelectedItems(0).SubItems(1).Text
            Process.Start("explorer.exe", dir)
        End With
    End Sub
    Private Sub Frm_Act_Documentos_Digitales_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            'Frm_Principal.TV.Enabled = True
            'Frm_Principal.TV.SelectedNode = Frm_Principal.TV.Nodes(0)
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Dim NOMBRE_ARCHIVO As String
    Dim NOMBRES_APELLIDOS As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        MENSAJE = MsgBox("Se dispone a Actualizar los nombres de los Archivos Digitales incorporando los codigos de Empleado , ¿Esta seguro de Continuar?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            Application.DoEvents()
            For i As Integer = 0 To lvFics.Items.Count - 1
                NOMBRE_ARCHIVO = Me.lvFics.Items(i).Text
                Me.LabelInfo.Text = NOMBRE_ARCHIVO
                NOMBRES_APELLIDOS = Mid(NOMBRE_ARCHIVO, 1, Len(NOMBRE_ARCHIVO) - 4)
                If InStr(NOMBRES_APELLIDOS, "-") Then
                    CODIGO_EMPLEADO = Mid(NOMBRES_APELLIDOS, 1, 10)
                    Call BUSCAR_EN_VISTA_MAESTRO_DE_CARGOS_SIT_2_a()
                Else
                    Call BUSCAR_EN_VISTA_MAESTRO_DE_CARGOS_SIT_2()
                End If
                If CODIGO_EMPLEADO <> "" Then
                    Call VERIFICAR_SI_ARCHIVO_EXISTE
                    If ARCHIVO_EXISTE = False Then
                        Call RENOMBRAR_ARCHIVO()
                    End If
                End If
                Me.LabelInfo.Refresh()
            Next
            '--------------------------------------------------------------
            Static yaEstoy As Boolean
            If yaEstoy Then
                cancelar = True
                Me.btnBuscar.Text = "Cancelando..."
                Application.DoEvents()
                Exit Sub
            End If
            yaEstoy = True
            With My.Settings
                .Location = Me.Location
                .Size = Me.Size
                .Directorio = Me.txtDir.Text
                .Filtro = Me.txtFiltro.Text
                .conSubDir = Me.chkConSubDir.Checked
                .IgnorarError = Me.chkIgnorarError.Checked
                .Save()
                Dim di As New DirectoryInfo(.Directorio)
                Me.lvFics.Items.Clear()
                Me.LabelInfo.Text = "Buscando los ficheros..."
                Me.Cursor = Cursors.AppStarting
                Me.btnBuscar.Text = "Buscando..."
                Me.Refresh()
                recorrerDir(di)
            End With
            Me.Cursor = Cursors.Default
            Me.LabelInfo.Text = "Se han hallado " & Me.lvFics.Items.Count & " ficheros"
            Me.btnBuscar.Text = "Buscar"
            If Me.lvFics.Items.Count > 0 Then
                Me.btnAbrirDir.Enabled = True
                Me.btnAbrirFic.Enabled = True
            End If
            Me.Refresh()
            cancelar = False
            yaEstoy = False
            '--------------------------------------------------------------
        End If
    End Sub
    Dim ARCHIVO_EXISTE As Boolean
    Private Sub VERIFICAR_SI_ARCHIVO_EXISTE()
        ARCHIVO_EXISTE = My.Computer.FileSystem.FileExists(Me.txtDir.Text & "\" & CODIGO_EMPLEADO & "-" & NOMBRES_APELLIDOS_CORRECTO & ".pdf")
    End Sub
    Dim CODIGO_EMPLEADO As String
    Dim NOMBRES_APELLIDOS_CORRECTO As String
    Private Sub BUSCAR_EN_VISTA_MAESTRO_DE_CARGOS_SIT_2_a()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [VISTA MAESTRO DE CARGOS SIT 2] WHERE CODIGO = '" & Trim(CODIGO_EMPLEADO) & "' ORDER BY NA", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                CODIGO_EMPLEADO = MiDataTable.Rows(0).Item(22).ToString
                NOMBRES_APELLIDOS_CORRECTO = MiDataTable.Rows(0).Item(23).ToString
            Else
                CODIGO_EMPLEADO = ""
                NOMBRES_APELLIDOS_CORRECTO = ""
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub BUSCAR_EN_VISTA_MAESTRO_DE_CARGOS_SIT_2()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [VISTA MAESTRO DE CARGOS SIT 2] WHERE NA = '" & Trim(NOMBRES_APELLIDOS) & "' ORDER BY NA", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                CODIGO_EMPLEADO = MiDataTable.Rows(0).Item(40).ToString
                NOMBRES_APELLIDOS_CORRECTO = MiDataTable.Rows(0).Item(41).ToString
            Else
                CODIGO_EMPLEADO = ""
                NOMBRES_APELLIDOS_CORRECTO = ""
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub RENOMBRAR_ARCHIVO()
        My.Computer.FileSystem.RenameFile(txtDir.Text & "\" & NOMBRE_ARCHIVO, CODIGO_EMPLEADO & "-" & NOMBRES_APELLIDOS_CORRECTO & ".pdf")
    End Sub
    Dim ID_EMPLEADO As Integer
    Dim ARCHIVO_DE_DESTINO As String
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Tipo de Documento Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        MENSAJE = MsgBox("Se dispone a Ingresar los Archivos Digitales con la Clasificación <" & Me.ComboBox1.Text & "> al Sistema, ¿Esta seguro de Continuar?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            Application.DoEvents()
            Me.Cursor = Cursors.WaitCursor
            Call BUSCAR_DIRECTORIO()
            For i As Integer = 0 To lvFics.Items.Count - 1
                NOMBRE_ARCHIVO = Me.lvFics.Items(i).Text
                Me.LabelInfo.Text = NOMBRE_ARCHIVO
                CODIGO_EMPLEADO = Mid(NOMBRE_ARCHIVO, 1, 10)
                Call BUSCAR_ID_EMPLEADO()
                If ID_EMPLEADO <> 0 Then
                    Me.Text = "Actualizar - Documentos Digitales [" & CODIGO_EMPLEADO & "]"
                    Call GENERAR_CODIGO_ID()
                    ARCHIVO_DE_DESTINO = CODIGO_EMPLEADO & "-" & CODIGO_ID_DOCUMENTO & "-" & Me.ComboBox1.Text & ".pdf"
                    Call INGRESAR_REGISTRO()
                    Call COPIAR_ARCHIVO_SERVIDOR()
                    Call RENOMBRAR_ARCHIVO_SERVIDOR()
                End If
                Me.LabelInfo.Refresh()
            Next
            Me.Cursor = Cursors.Default
            MsgBox("El proceso a Finalizado Correctamente", vbInformation, "Mensaje del Sistema")
        End If
    End Sub
    Private Sub COPIAR_ARCHIVO_SERVIDOR()
        My.Computer.FileSystem.CopyFile(txtDir.Text & "\" & NOMBRE_ARCHIVO, dDIRECTORIO & "\" & NOMBRE_ARCHIVO)
    End Sub
    Private Sub RENOMBRAR_ARCHIVO_SERVIDOR()
        My.Computer.FileSystem.RenameFile(dDIRECTORIO & "\" & NOMBRE_ARCHIVO, ARCHIVO_DE_DESTINO)
    End Sub
    Private Sub INGRESAR_REGISTRO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE DOCUMENTOS DIGITALES] (ID_M_DD, ID_M_P, ID_D1, ID_T_DOC, NOMBRE, OBSERVACION)" &
                                  "values (" & CInt(CODIGO_ID_DOCUMENTO) & ", " & ID_EMPLEADO & ", 1, " & Me.ComboBox1.SelectedValue & ", '" & ARCHIVO_DE_DESTINO & "', 'INGRESADO DESDE EL MODULO DE ACTUALIZACION DE DOCUMENTOS MASIVO')"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
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
    Dim CODIGO_ID_DOCUMENTO As Integer
    Private Sub GENERAR_CODIGO_ID()
        CODIGO_ID_DOCUMENTO = 0
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
            CODIGO_ID_DOCUMENTO = CODIGO
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub BUSCAR_ID_EMPLEADO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT ID_M_P FROM [VISTA MAESTRO DE CARGOS SIT 2] WHERE CODIGO = '" & Trim(CODIGO_EMPLEADO) & "' ORDER BY NA", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                ID_EMPLEADO = MiDataTable.Rows(0).Item(0).ToString
            Else
                ID_EMPLEADO = 0
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub lvFics_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lvFics.SelectedIndexChanged

    End Sub
End Class