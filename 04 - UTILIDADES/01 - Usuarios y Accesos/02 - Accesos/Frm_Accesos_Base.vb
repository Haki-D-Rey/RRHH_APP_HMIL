Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Accesos_Base
    Dim aIndex(20) As Integer
    Dim nLevel As Integer
    Dim oNode As TreeNode
    Private Sub Frm_Accesos_Base_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Accesos_Base_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Cursor = Cursors.WaitCursor
        Call CARGAR_TV()
        TV.Nodes(0).ForeColor = Color.DarkGreen : TV.Nodes(0).Expand()
        TV.Nodes(0).Nodes(0).ForeColor = Color.Blue : TV.Nodes(0).Nodes(0).Expand()
        TV.Nodes(0).Nodes(1).ForeColor = Color.Blue : TV.Nodes(0).Nodes(1).Expand()
        TV.Nodes(0).Nodes(2).ForeColor = Color.Blue : TV.Nodes(0).Nodes(2).Expand()
        TV.Nodes(0).Nodes(3).ForeColor = Color.Blue : TV.Nodes(0).Nodes(3).Expand()
        TV.Nodes(0).Nodes(4).ForeColor = Color.Blue : TV.Nodes(0).Nodes(4).Expand()
        TV.Nodes(0).Nodes(5).ForeColor = Color.Blue : TV.Nodes(0).Nodes(5).Expand()
        TV.Nodes(0).Nodes(5).Nodes(2).ForeColor = Color.Red
        TV.SelectedNode = TV.Nodes(0)
        Me.Cursor = Cursors.Default
        Call RECORRER_Y_CARGAR_NODOS()
        TV.ExpandAll()
        TV.SelectedNode = TV.Nodes(0)
        NODO_TAG = 0
        NODO_TEXT = ""
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub CARGAR_TV()
        TV.BeginUpdate()
        TV.Nodes.Clear()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [MAESTRO DE NODOS BASE] WHERE ID_NODO_PADRE = 0 ", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                Dim A As Integer
                Dim dr As DataRow
                For Each dr In MiDataTable.Rows
                    aIndex(nLevel) = A
                    oNode = New TreeNode
                    oNode.Text = "[" & dr("CLAVE") & "] " & dr("DESCRIPCION")
                    oNode.Tag &= dr("ID_NODO")
                    oNode.ImageIndex = 0
                    TV.Nodes.Add(oNode)
                    BUSCAR_ITEMS(dr("ID_NODO"))
                    A += 1
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
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [MAESTRO DE NODOS BASE] WHERE ID_NODO_PADRE = " & nId & " order by ID_NODO, ID_NODO_PADRE", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                Dim oNode As TreeNode
                Dim B As Integer
                Dim dr As DataRow
                For Each dr In MiDataTable.Rows
                    aIndex(nLevel) = B
                    oNode = New TreeNode
                    oNode.Text = "[" & dr("CLAVE") & "] " & dr("DESCRIPCION")
                    oNode.Tag &= dr("ID_NODO")
                    oNode.Expand()
                    Select Case nLevel
                        Case 1
                            TV.Nodes(aIndex(0)).Nodes.Add(oNode)
                            oNode.ImageIndex = 3
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
                    B += 1
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
    Dim tn As TreeNode
    Private Sub RECORRER_NODOS_SELECCIONADOS(treeNode As TreeNode)
        Try
            For Each tn In treeNode.Nodes
                If tn.Checked = True Then
                    Call ACTUALIZAR_TABLA()
                End If
                If tn.Checked = False Then
                    Call ACTUALIZAR_TABLA_1()
                End If
                RECORRER_NODOS_SELECCIONADOS(tn)
            Next
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        End Try
    End Sub
    Private Sub RECORRER_NODOS()
        Dim nodes As TreeNodeCollection = TV.Nodes
        For Each n As TreeNode In nodes
            RECORRER_NODOS_SELECCIONADOS(n)
        Next
    End Sub
    Private Sub ACTUALIZAR_TABLA_1()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE NODOS BASE] SET ACCESO = '0' WHERE ID_NODO = " & tn.Tag & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ACTUALIZAR_TABLA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE NODOS BASE] SET ACCESO = '1' WHERE ID_NODO = " & tn.Tag & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    '*********************************************************************
    Private Sub RECORRER_Y_CARGAR_NODOS_SELECCIONADOS(treeNode As TreeNode)
        Try
            For Each tn In treeNode.Nodes
                'MessageBox.Show(tn.Tag & " " & tn.Text)
                Call VERIFICAR_ESTADO_DE_NODO_EN_LA_TABLA()
                If ESTADO_DE_NODO_EN_LA_TABLA = True Then
                    tn.Checked = True
                Else
                    tn.Checked = False
                End If
                RECORRER_Y_CARGAR_NODOS_SELECCIONADOS(tn)
            Next
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        End Try
    End Sub
    Dim xn1 As TreeNode
    Private Sub RECORRER_Y_CARGAR_NODOS()
        Dim nodes As TreeNodeCollection = TV.Nodes
        For Each n As TreeNode In nodes
            xn1 = n
            RECORRER_Y_CARGAR_NODOS_SELECCIONADOS(n)
            Call VERIFICAR_ESTADO_DE_NODO_EN_LA_TABLA_1()
            If ESTADO_DE_NODO_EN_LA_TABLA_1 = True Then
                n.Checked = True
            Else
                n.Checked = False
            End If
        Next
    End Sub
    Dim ESTADO_DE_NODO_EN_LA_TABLA_1 As Boolean
    Private Sub VERIFICAR_ESTADO_DE_NODO_EN_LA_TABLA_1()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [MAESTRO DE NODOS BASE] WHERE ID_NODO = " & xn1.Tag & " AND ACCESO = '1'", Conexion.CxRRHH)

            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                ESTADO_DE_NODO_EN_LA_TABLA_1 = True
            Else
                ESTADO_DE_NODO_EN_LA_TABLA_1 = False
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim ESTADO_DE_NODO_EN_LA_TABLA As Boolean
    Private Sub VERIFICAR_ESTADO_DE_NODO_EN_LA_TABLA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [MAESTRO DE NODOS BASE] WHERE ID_NODO = " & tn.Tag & " AND ACCESO = '1'", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                ESTADO_DE_NODO_EN_LA_TABLA = True
            Else
                ESTADO_DE_NODO_EN_LA_TABLA = False
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
        Me.Close()
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ''Call VERIFICAR_ACCESOS("037") : If HAY_ACCESO = False Then : Exit Sub : End If
        Me.Cursor = Cursors.WaitCursor
        Call RECORRER_NODOS()
        Me.Close()
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub TV_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TV.AfterSelect
        NODO_TAG = e.Node.Tag
        NODO_TEXT = e.Node.Text
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ''Call VERIFICAR_ACCESOS("035") : If HAY_ACCESO = False Then : Exit Sub : End If
        If NODO_TAG = 0 Then
            MsgBox("Debe seleccionar la Ubicación donde se agregara el Nuevo Nodo", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        Frm_Accesos_Base_Add.ShowDialog()
        If CERRAR = False Then
            Me.Cursor = Cursors.WaitCursor
            Call CARGAR_TV()
            TV.Nodes(0).ForeColor = Color.DarkGreen : TV.Nodes(0).Expand()
            TV.Nodes(0).Nodes(0).ForeColor = Color.Blue : TV.Nodes(0).Nodes(0).Expand()
            TV.Nodes(0).Nodes(1).ForeColor = Color.Blue : TV.Nodes(0).Nodes(1).Expand()
            TV.Nodes(0).Nodes(2).ForeColor = Color.Blue : TV.Nodes(0).Nodes(2).Expand()
            TV.Nodes(0).Nodes(3).ForeColor = Color.Blue : TV.Nodes(0).Nodes(3).Expand()
            TV.Nodes(0).Nodes(4).ForeColor = Color.Blue : TV.Nodes(0).Nodes(4).Expand()
            TV.Nodes(0).Nodes(5).ForeColor = Color.Blue : TV.Nodes(0).Nodes(5).Expand()
            TV.Nodes(0).Nodes(5).Nodes(2).ForeColor = Color.Red
            TV.SelectedNode = TV.Nodes(0)
            Me.Cursor = Cursors.Default
            Call RECORRER_Y_CARGAR_NODOS()
            TV.ExpandAll()
            TV.SelectedNode = TV.Nodes(0)
            NODO_TAG = 0
            NODO_TEXT = ""
            Me.Cursor = Cursors.Default
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ''Call VERIFICAR_ACCESOS("036") : If HAY_ACCESO = False Then : Exit Sub : End If
        If NODO_TAG = 0 Then
            MsgBox("Debe seleccionar el Nodo que desea Eliminar, recuerde que si el Nodo que intenta Eliminar tiene nodos internos estos deben ser eliminados primeramente", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        MENSAJE = MsgBox("¿Esta seguro de Eliminar el registro <" & NODO_TEXT & "> seleccionado?, recuerde que si el Nodo que intenta Eliminar tiene nodos internos estos deben ser eliminados primeramente", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            Call ELIMINAR_NODO_EN_MAESTRO_DE_NODOS_BASE()
            Call ELIMINAR_NODO_EN_MAESTRO_DE_NODOS_USUARIOS()
            Me.Cursor = Cursors.WaitCursor
            Call CARGAR_TV()
            TV.Nodes(0).ForeColor = Color.DarkGreen : TV.Nodes(0).Expand()
            TV.Nodes(0).Nodes(0).ForeColor = Color.Blue : TV.Nodes(0).Nodes(0).Expand()
            TV.Nodes(0).Nodes(1).ForeColor = Color.Blue : TV.Nodes(0).Nodes(1).Expand()
            TV.Nodes(0).Nodes(2).ForeColor = Color.Blue : TV.Nodes(0).Nodes(2).Expand()
            TV.Nodes(0).Nodes(3).ForeColor = Color.Blue : TV.Nodes(0).Nodes(3).Expand()
            TV.Nodes(0).Nodes(4).ForeColor = Color.Blue : TV.Nodes(0).Nodes(4).Expand()
            TV.Nodes(0).Nodes(5).ForeColor = Color.Blue : TV.Nodes(0).Nodes(5).Expand()
            TV.Nodes(0).Nodes(5).Nodes(2).ForeColor = Color.Red
            TV.SelectedNode = TV.Nodes(0)
            Me.Cursor = Cursors.Default
            Call RECORRER_Y_CARGAR_NODOS()
            TV.ExpandAll()
            TV.SelectedNode = TV.Nodes(0)
            NODO_TAG = 0
            NODO_TEXT = ""
            Me.Cursor = Cursors.Default
        End If
    End Sub
    Private Sub ELIMINAR_NODO_EN_MAESTRO_DE_NODOS_USUARIOS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE NODOS USUARIOS] WHERE ID_NODO = " & CInt(NODO_TAG) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ELIMINAR_NODO_EN_MAESTRO_DE_NODOS_BASE()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "DELETE FROM [MAESTRO DE NODOS BASE] WHERE ID_NODO = " & CInt(NODO_TAG) & ""
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