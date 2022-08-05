Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Catalogos_Seleccion
    Private Sub Frm_Catalogos_Seleccion_Seleccion_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        CATALOGOS = False
        'Frm_Principal.Timer1.Start()
    End Sub
    Private Sub Frm_Catalogos_Seleccion_Seleccion_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            CATALOGOS = False
            'Frm_Principal.Timer1.Start()
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Catalogos_Seleccion_Seleccion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CATALOGOS = True
        'Frm_Principal.Timer1.Stop()
        Me.Cursor = Cursors.WaitCursor
        Call CARGAR_TV()
        TV.ExpandAll()
        TV.SelectedNode = TV.Nodes(0)
        Me.Height = Frm_Principal.Size.Height - 10
        Me.Cursor = Cursors.Default
    End Sub
    Dim aIndex(20) As Integer
    Dim nLevel As Integer
    Dim oNode As TreeNode
    Private Sub CARGAR_TV()
        TV.BeginUpdate()
        TV.Nodes.Clear()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [MAESTRO DE NODOS CATALOGOS] WHERE ID_NODO_PADRE = 0 ", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                Dim i As Integer
                Dim dr As DataRow
                For Each dr In MiDataTable.Rows
                    aIndex(nLevel) = i
                    oNode = New TreeNode
                    oNode.Text = dr("DESCRIPCION")
                    oNode.Tag &= dr("ID_NODO")
                    oNode.ImageIndex = 1
                    TV.Nodes.Add(oNode)
                    BUSCAR_ITEMS(dr("ID_NODO"))
                    i += 1
                Next
            End If
            TV.EndUpdate()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub BUSCAR_ITEMS(ByVal nId As Integer)
        nLevel += 1
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [MAESTRO DE NODOS CATALOGOS] WHERE ID_NODO_PADRE = " & nId & " ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                Dim oNode As TreeNode
                Dim i As Integer
                Dim dr As DataRow
                For Each dr In MiDataTable.Rows
                    aIndex(nLevel) = i
                    oNode = New TreeNode
                    oNode.Text = dr("DESCRIPCION")
                    oNode.Tag &= dr("ID_NODO")
                    oNode.Expand()
                    Select Case nLevel
                        Case 1
                            TV.Nodes(aIndex(0)).Nodes.Add(oNode)
                            oNode.ImageIndex = 2
                        Case 2
                            TV.Nodes(aIndex(0)).Nodes(aIndex(1)).Nodes.Add(oNode)
                            oNode.ImageIndex = 3
                        Case 3
                            TV.Nodes(aIndex(0)).Nodes(aIndex(1)).Nodes(aIndex(2)).Nodes.Add(oNode)
                            oNode.ImageIndex = 3
                        Case 4
                            TV.Nodes(aIndex(0)).Nodes(aIndex(1)).Nodes(aIndex(2)).Nodes(aIndex(3)).Nodes.Add(oNode)
                            oNode.ImageIndex = 3
                        Case 5
                            TV.Nodes(aIndex(0)).Nodes(aIndex(1)).Nodes(aIndex(2)).Nodes(aIndex(3)).Nodes(aIndex(4)).Nodes.Add(oNode)
                            oNode.ImageIndex = 3
                        Case 6
                            TV.Nodes(aIndex(0)).Nodes(aIndex(1)).Nodes(aIndex(2)).Nodes(aIndex(3)).Nodes(aIndex(4)).Nodes(aIndex(5)).Nodes.Add(oNode)
                            oNode.ImageIndex = 3
                        Case 7
                            TV.Nodes(aIndex(0)).Nodes(aIndex(1)).Nodes(aIndex(2)).Nodes(aIndex(3)).Nodes(aIndex(4)).Nodes(aIndex(5)).Nodes(aIndex(6)).Nodes.Add(oNode)
                            oNode.ImageIndex = 3
                        Case 8
                            TV.Nodes(aIndex(0)).Nodes(aIndex(1)).Nodes(aIndex(2)).Nodes(aIndex(3)).Nodes(aIndex(4)).Nodes(aIndex(5)).Nodes(aIndex(6)).Nodes(aIndex(7)).Nodes.Add(oNode)
                            oNode.ImageIndex = 3
                        Case 9
                            TV.Nodes(aIndex(0)).Nodes(aIndex(1)).Nodes(aIndex(2)).Nodes(aIndex(3)).Nodes(aIndex(4)).Nodes(aIndex(5)).Nodes(aIndex(6)).Nodes(aIndex(7)).Nodes(aIndex(8)).Nodes.Add(oNode)
                            oNode.ImageIndex = 3
                        Case 10
                            TV.Nodes(aIndex(0)).Nodes(aIndex(1)).Nodes(aIndex(2)).Nodes(aIndex(3)).Nodes(aIndex(4)).Nodes(aIndex(5)).Nodes(aIndex(6)).Nodes(aIndex(7)).Nodes(aIndex(8)).Nodes(aIndex(9)).Nodes.Add(oNode)
                            oNode.ImageIndex = 3
                        Case 11
                            TV.Nodes(aIndex(0)).Nodes(aIndex(1)).Nodes(aIndex(2)).Nodes(aIndex(3)).Nodes(aIndex(4)).Nodes(aIndex(5)).Nodes(aIndex(6)).Nodes(aIndex(7)).Nodes(aIndex(8)).Nodes(aIndex(9)).Nodes(aIndex(10)).Nodes.Add(oNode)
                            oNode.ImageIndex = 3
                        Case 12
                            TV.Nodes(aIndex(0)).Nodes(aIndex(1)).Nodes(aIndex(2)).Nodes(aIndex(3)).Nodes(aIndex(4)).Nodes(aIndex(5)).Nodes(aIndex(6)).Nodes(aIndex(7)).Nodes(aIndex(8)).Nodes(aIndex(9)).Nodes(aIndex(10)).Nodes(aIndex(11)).Nodes.Add(oNode)
                            oNode.ImageIndex = 3
                        Case 13
                            TV.Nodes(aIndex(0)).Nodes(aIndex(1)).Nodes(aIndex(2)).Nodes(aIndex(3)).Nodes(aIndex(4)).Nodes(aIndex(5)).Nodes(aIndex(6)).Nodes(aIndex(7)).Nodes(aIndex(8)).Nodes(aIndex(9)).Nodes(aIndex(10)).Nodes(aIndex(11)).Nodes(aIndex(12)).Nodes.Add(oNode)
                            oNode.ImageIndex = 3
                        Case 14
                            TV.Nodes(aIndex(0)).Nodes(aIndex(1)).Nodes(aIndex(2)).Nodes(aIndex(3)).Nodes(aIndex(4)).Nodes(aIndex(5)).Nodes(aIndex(6)).Nodes(aIndex(7)).Nodes(aIndex(8)).Nodes(aIndex(9)).Nodes(aIndex(10)).Nodes(aIndex(11)).Nodes(aIndex(12)).Nodes(aIndex(13)).Nodes.Add(oNode)
                            oNode.ImageIndex = 3
                        Case 15
                            TV.Nodes(aIndex(0)).Nodes(aIndex(1)).Nodes(aIndex(2)).Nodes(aIndex(3)).Nodes(aIndex(4)).Nodes(aIndex(5)).Nodes(aIndex(6)).Nodes(aIndex(7)).Nodes(aIndex(8)).Nodes(aIndex(9)).Nodes(aIndex(10)).Nodes(aIndex(11)).Nodes(aIndex(12)).Nodes(aIndex(13)).Nodes(aIndex(14)).Nodes.Add(oNode)
                            oNode.ImageIndex = 3
                        Case 16
                            TV.Nodes(aIndex(0)).Nodes(aIndex(1)).Nodes(aIndex(2)).Nodes(aIndex(3)).Nodes(aIndex(4)).Nodes(aIndex(5)).Nodes(aIndex(6)).Nodes(aIndex(7)).Nodes(aIndex(8)).Nodes(aIndex(9)).Nodes(aIndex(10)).Nodes(aIndex(11)).Nodes(aIndex(12)).Nodes(aIndex(13)).Nodes(aIndex(14)).Nodes(aIndex(15)).Nodes.Add(oNode)
                            oNode.ImageIndex = 3
                        Case 17
                            TV.Nodes(aIndex(0)).Nodes(aIndex(1)).Nodes(aIndex(2)).Nodes(aIndex(3)).Nodes(aIndex(4)).Nodes(aIndex(5)).Nodes(aIndex(6)).Nodes(aIndex(7)).Nodes(aIndex(8)).Nodes(aIndex(9)).Nodes(aIndex(10)).Nodes(aIndex(11)).Nodes(aIndex(12)).Nodes(aIndex(13)).Nodes(aIndex(14)).Nodes(aIndex(15)).Nodes(aIndex(16)).Nodes.Add(oNode)
                            oNode.ImageIndex = 3
                        Case 18
                            TV.Nodes(aIndex(0)).Nodes(aIndex(1)).Nodes(aIndex(2)).Nodes(aIndex(3)).Nodes(aIndex(4)).Nodes(aIndex(5)).Nodes(aIndex(6)).Nodes(aIndex(7)).Nodes(aIndex(8)).Nodes(aIndex(9)).Nodes(aIndex(10)).Nodes(aIndex(11)).Nodes(aIndex(12)).Nodes(aIndex(13)).Nodes(aIndex(14)).Nodes(aIndex(15)).Nodes(aIndex(16)).Nodes(aIndex(17)).Nodes.Add(oNode)
                            oNode.ImageIndex = 3
                        Case 19
                            TV.Nodes(aIndex(0)).Nodes(aIndex(1)).Nodes(aIndex(2)).Nodes(aIndex(3)).Nodes(aIndex(4)).Nodes(aIndex(5)).Nodes(aIndex(6)).Nodes(aIndex(7)).Nodes(aIndex(8)).Nodes(aIndex(9)).Nodes(aIndex(10)).Nodes(aIndex(11)).Nodes(aIndex(12)).Nodes(aIndex(13)).Nodes(aIndex(14)).Nodes(aIndex(15)).Nodes(aIndex(16)).Nodes(aIndex(17)).Nodes(aIndex(18)).Nodes.Add(oNode)
                            oNode.ImageIndex = 3
                        Case 20
                            TV.Nodes(aIndex(0)).Nodes(aIndex(1)).Nodes(aIndex(2)).Nodes(aIndex(3)).Nodes(aIndex(4)).Nodes(aIndex(5)).Nodes(aIndex(6)).Nodes(aIndex(7)).Nodes(aIndex(8)).Nodes(aIndex(9)).Nodes(aIndex(10)).Nodes(aIndex(11)).Nodes(aIndex(12)).Nodes(aIndex(13)).Nodes(aIndex(14)).Nodes(aIndex(15)).Nodes(aIndex(16)).Nodes(aIndex(17)).Nodes(aIndex(18)).Nodes(aIndex(19)).Nodes.Add(oNode)
                            oNode.ImageIndex = 3
                    End Select
                    'Hago la recursividad para llenar mis datos de la tabla en n niveles llamando la funcion
                    BUSCAR_ITEMS(dr("ID_NODO"))
                    i += 1
                Next
            End If
            nLevel -= 1
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub TV_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TV.AfterSelect
        If e.Node.Text = "Tipo de Estructuras" Then
            Call VERIFICAR_ACCESOS("290") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_T_Estructura.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Estructura Nivel 1" Then
            Call VERIFICAR_ACCESOS("248") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Estructura_N1.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Estructura Nivel 2" Then
            Call VERIFICAR_ACCESOS("249") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Estructura_N2.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Estructura Nivel 3" Then
            Call VERIFICAR_ACCESOS("250") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Estructura_N3.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Estructura Nivel 4" Then
            Call VERIFICAR_ACCESOS("251") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Estructura_N4.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Estructura Nivel 5" Then
            Call VERIFICAR_ACCESOS("252") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Estructura_N5.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Estructura Nivel 6" Then
            Call VERIFICAR_ACCESOS("253") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Estructura_N6.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Estructura Nivel 7" Then
            Call VERIFICAR_ACCESOS("254") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Estructura_N7.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Estructura Nivel 8" Then
            Call VERIFICAR_ACCESOS("255") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Estructura_N8.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Cargos" Then
            Call VERIFICAR_ACCESOS("229") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Cargos.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Especialidades (Medicas y Otras)" Then
            Call VERIFICAR_ACCESOS("245") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Espe_Medicas.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Grado Plantilla" Then
            Call VERIFICAR_ACCESOS("263") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_GP.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Grado Real" Then
            Call VERIFICAR_ACCESOS("264") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_GR.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Nivel Academico" Then
            Call VERIFICAR_ACCESOS("274") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Nivel_Aca.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Nivel Profesional" Then
            Call VERIFICAR_ACCESOS("277") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Nivel_Prof.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Nacionalidad (Nacimiento)" Then
            Call VERIFICAR_ACCESOS("272") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Nacionalidad_N.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Nacionalidad (Domiciliar)" Then
            Call VERIFICAR_ACCESOS("271") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Nacionalidad_D.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Departamento (Nacimiento)" Then
            Call VERIFICAR_ACCESOS("242") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Departamento_N.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Departamento (Domiciliar)" Then
            Call VERIFICAR_ACCESOS("241") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Departamento_D.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Municipio (Nacimiento)" Then
            Call VERIFICAR_ACCESOS("270") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Municipio_N.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Municipio (Domiciliar)" Then
            Call VERIFICAR_ACCESOS("269") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Municipio_D.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Estado Civil" Then
            Call VERIFICAR_ACCESOS("247") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Estado_C.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Tipo de Horario" Then
            Call VERIFICAR_ACCESOS("294") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_T_Horario.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Tipo de Afiliacion" Then
            Call VERIFICAR_ACCESOS("288") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_T_Afiliacion.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Tipo de Plaza" Then
            Call VERIFICAR_ACCESOS("295") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_T_Plaza.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Categoria Salarial" Then
            Call VERIFICAR_ACCESOS("231") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Categoria_S.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Parentela" Then
            Call VERIFICAR_ACCESOS("281") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Parentela.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Tipo de Estudio" Then
            Call VERIFICAR_ACCESOS("291") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_T_Estudio.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Habilidades Administrativas" Then
            Call VERIFICAR_ACCESOS("265") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Habilidades_A.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Idiomas" Then
            Call VERIFICAR_ACCESOS("266") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Idiomas.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Nominas" Then
            Call VERIFICAR_ACCESOS("278") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Nominas.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Movimientos en Nominas" Then
            Call VERIFICAR_ACCESOS("268") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Movimientos_Nomina.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Eventualidades" Then
            Call VERIFICAR_ACCESOS("256") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Eventualidades.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Curva y Talla" Then
            Call VERIFICAR_ACCESOS("240") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_CyT.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Clasificacion de Personal" Then
            Call VERIFICAR_ACCESOS("236") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_C_Personal.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Clasificacion de Condecoraciones" Then
            Call VERIFICAR_ACCESOS("235") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_C_Condecoraciones.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Condecoraciones" Then
            Call VERIFICAR_ACCESOS("237") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Condecoraciones.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Causal de Movimiento Real" Then
            Call VERIFICAR_ACCESOS("233") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Causas_Reales.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Causal de Movimiento Legal" Then
            Call VERIFICAR_ACCESOS("232") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Causas_Legales.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Pais" Then
            Call VERIFICAR_ACCESOS("280") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Pais.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Tipo Sanguineo" Then
            Call VERIFICAR_ACCESOS("299") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Tipo_Sanguineo.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Tipo Documento" Then
            Call VERIFICAR_ACCESOS("298") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Tipo_Documento.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Tipo de Anexos" Then
            Call VERIFICAR_ACCESOS("289") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Tipo_de_Anexos.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Anexos" Then
            Call VERIFICAR_ACCESOS("228") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Anexos.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Categoria de Licencia de Conducir" Then
            Call VERIFICAR_ACCESOS("230") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Categoria_de_Lic.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Ipss" Then
            Call VERIFICAR_ACCESOS("267") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Ipss.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Tipo de Situacion de Personal" Then
            Call VERIFICAR_ACCESOS("297") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Tipo_de_Sit_Personal.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Situacion de Personal" Then
            Call VERIFICAR_ACCESOS("286") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Sit_Personal.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Funcionalidad" Then
            Call VERIFICAR_ACCESOS("262") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Funcionalidad.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Contratos y Evaluaciones" Then
            Call VERIFICAR_ACCESOS("239") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Contratos_y_Evaluaciones.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Tipo de Accidente" Then
            Call VERIFICAR_ACCESOS("287") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Tipo_de_Accidente.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Agente Material General" Then
            Call VERIFICAR_ACCESOS("226") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Agente_Material_General.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Agente Material Sub Estandar" Then
            Call VERIFICAR_ACCESOS("227") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Agente_Material_Sub_Estandar.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Forma de Ocurrencia" Then
            Call VERIFICAR_ACCESOS("261") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Forma_de_Ocurrencia.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Parte del Cuerpo Lesionada" Then
            Call VERIFICAR_ACCESOS("282") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Parte_del_Cuerpo_Lesionada.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Naturaleza de la Lesion" Then
            Call VERIFICAR_ACCESOS("273") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Naturaleza_de_la_Lesion.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Organizativas - Sistema de Prevencion" Then
            Call VERIFICAR_ACCESOS("279") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Organizativas_Sistema_de_Prevencion.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Factores Personales" Then
            Call VERIFICAR_ACCESOS("259") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Factores_Personales.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Factores de Trabajo" Then
            Call VERIFICAR_ACCESOS("258") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Factores_de_Trabajo.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Actos Inseguros" Then
            Call VERIFICAR_ACCESOS("224") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Actos_Inseguros.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Condiciones Inseguras" Then
            Call VERIFICAR_ACCESOS("238") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Condiciones_Inseguras.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Factores de Riesgo" Then
            Call VERIFICAR_ACCESOS("257") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Factores_de_Riesgo.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Nivel de Gravedad" Then
            Call VERIFICAR_ACCESOS("276") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Nivel_de_Gravedad.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Acciones de Mejora Continua" Then
            Call VERIFICAR_ACCESOS("223") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Acciones_de_Mejora_Continua.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Periodo Dosimetro" Then
            Call VERIFICAR_ACCESOS("283") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Periodo_Dosimetro.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Empresas Tercerizadas" Then
            Call VERIFICAR_ACCESOS("244") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_E_Tercerizadas.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Tipo de Evaluacion" Then
            Call VERIFICAR_ACCESOS("292") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_T_Evaluacion.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Tipo de Examenes" Then
            Call VERIFICAR_ACCESOS("293") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_T_Examenes.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Resultados de Evaluacion" Then
            Call VERIFICAR_ACCESOS("285") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_R_Evaluacion.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Tipo de Resultado de Examen" Then
            Call VERIFICAR_ACCESOS("296") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_T_R_Examen.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Causas Reales de Eventos" Then
            Call VERIFICAR_ACCESOS("234") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Causas_R_Eventos.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Destinatarios" Then
            Call VERIFICAR_ACCESOS("243") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Destinatarios.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Feriados" Then
            Call VERIFICAR_ACCESOS("260") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Feriados.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Nivel de Competencia" Then
            Call VERIFICAR_ACCESOS("275") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Nivel_de_Competencia.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Establecimientos" Then
            Call VERIFICAR_ACCESOS("246") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Establecimientos.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Administradores de Recursos" Then
            Call VERIFICAR_ACCESOS("225") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Administradores_R.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Prefijos" Then
            Call VERIFICAR_ACCESOS("284") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Prefijos.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Inmunizaciones (Vacunas)" Then
            Call VERIFICAR_ACCESOS("312") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Vacunas.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        If e.Node.Text = "Inmunizaciones (Dosis)" Then
            Call VERIFICAR_ACCESOS("313") : If HAY_ACCESO = False Then : Exit Sub : End If
            Me.Cursor = Cursors.WaitCursor
            Frm_Cat_Dosis.ShowDialog()
            Me.TV.SelectedNode = Nothing
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
SALTO:
    End Sub
    Private Sub Frm_Catalogos_Seleccion_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        CATALOGOS = False
        Frm_Principal.Timer1.Start()
    End Sub
    Private Sub Frm_Catalogos_Seleccion_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        CATALOGOS = False
        Frm_Principal.Timer1.Start()
    End Sub
    Private Sub Frm_Catalogos_Seleccion_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        CATALOGOS = False
        Frm_Principal.Timer1.Start()
    End Sub
End Class