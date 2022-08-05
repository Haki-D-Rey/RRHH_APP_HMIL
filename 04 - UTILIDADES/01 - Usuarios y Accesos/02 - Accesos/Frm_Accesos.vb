Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Accesos
    Dim aIndex1(20) As Integer
    Dim nLevel1 As Integer
    Dim oNode1 As TreeNode
    Private Sub Frm_Accesos_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.Close()
        End If
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Me.Close()
    End Sub
    Private Sub Frm_Accesos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Accesos <" & uCUENTA & ">"
        Me.Cursor = Cursors.WaitCursor
        Call CARGAR_TV_1()
        TV1.Nodes(0).ForeColor = Color.DarkGreen : TV1.Nodes(0).Expand()
        TV1.Nodes(0).Nodes(0).ForeColor = Color.Blue : TV1.Nodes(0).Nodes(0).Expand()
        TV1.Nodes(0).Nodes(1).ForeColor = Color.Blue : TV1.Nodes(0).Nodes(1).Expand()
        TV1.Nodes(0).Nodes(2).ForeColor = Color.Blue : TV1.Nodes(0).Nodes(2).Expand()
        TV1.Nodes(0).Nodes(3).ForeColor = Color.Blue : TV1.Nodes(0).Nodes(3).Expand()
        TV1.Nodes(0).Nodes(4).ForeColor = Color.Blue : TV1.Nodes(0).Nodes(4).Expand()
        TV1.Nodes(0).Nodes(5).ForeColor = Color.Blue : TV1.Nodes(0).Nodes(5).Expand()
        TV1.Nodes(0).Nodes(5).Nodes(2).ForeColor = Color.Red
        TV1.SelectedNode = TV1.Nodes(0)
        Me.Cursor = Cursors.Default
        Call RECORRER_Y_CARGAR_NODOS_1()
        TV1.ExpandAll()
        TV1.SelectedNode = TV1.Nodes(0)
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub CARGAR_TV_1()
        TV1.BeginUpdate()
        TV1.Nodes.Clear()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [MAESTRO DE NODOS USUARIOS] WHERE ID_USUARIO = " & uID_USUARIO & " AND ID_NODO_PADRE = 0 ", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                Dim A As Integer
                Dim dr As DataRow
                For Each dr In MiDataTable.Rows
                    aIndex1(nLevel1) = A
                    oNode1 = New TreeNode
                    oNode1.Text = "[" & dr("CLAVE") & "] " & dr("DESCRIPCION")
                    oNode1.Tag &= dr("ID_NODO")
                    oNode1.ImageIndex = 0
                    TV1.Nodes.Add(oNode1)
                    BUSCAR_ITEMS(dr("ID_NODO"))
                    A += 1
                Next
            End If
            TV1.EndUpdate()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub BUSCAR_ITEMS(ByVal nId As Integer)
        nLevel1 += 1
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [MAESTRO DE NODOS USUARIOS] WHERE ID_USUARIO = " & uID_USUARIO & " AND ID_NODO_PADRE = " & nId & " order by ID_NODO, ID_NODO_PADRE", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                Dim oNode1 As TreeNode
                Dim B As Integer
                Dim dr As DataRow
                For Each dr In MiDataTable.Rows
                    aIndex1(nLevel1) = B
                    oNode1 = New TreeNode
                    oNode1.Text = "[" & dr("CLAVE") & "] " & dr("DESCRIPCION")
                    oNode1.Tag &= dr("ID_NODO")
                    oNode1.Expand()
                    Select Case nLevel1
                        Case 1
                            TV1.Nodes(aIndex1(0)).Nodes.Add(oNode1)
                            oNode1.ImageIndex = 3
                        Case 2
                            TV1.Nodes(aIndex1(0)).Nodes(aIndex1(1)).Nodes.Add(oNode1)
                            oNode1.ImageIndex = 3
                        Case 3
                            TV1.Nodes(aIndex1(0)).Nodes(aIndex1(1)).Nodes(aIndex1(2)).Nodes.Add(oNode1)
                            oNode1.ImageIndex = 3
                        Case 4
                            TV1.Nodes(aIndex1(0)).Nodes(aIndex1(1)).Nodes(aIndex1(2)).Nodes(aIndex1(3)).Nodes.Add(oNode1)
                            oNode1.ImageIndex = 3
                        Case 5
                            TV1.Nodes(aIndex1(0)).Nodes(aIndex1(1)).Nodes(aIndex1(2)).Nodes(aIndex1(3)).Nodes(aIndex1(4)).Nodes.Add(oNode1)
                            oNode1.ImageIndex = 3
                        Case 6
                            TV1.Nodes(aIndex1(0)).Nodes(aIndex1(1)).Nodes(aIndex1(2)).Nodes(aIndex1(3)).Nodes(aIndex1(4)).Nodes(aIndex1(5)).Nodes.Add(oNode1)
                            oNode1.ImageIndex = 3
                        Case 7
                            TV1.Nodes(aIndex1(0)).Nodes(aIndex1(1)).Nodes(aIndex1(2)).Nodes(aIndex1(3)).Nodes(aIndex1(4)).Nodes(aIndex1(5)).Nodes(aIndex1(6)).Nodes.Add(oNode1)
                            oNode1.ImageIndex = 3
                        Case 8
                            TV1.Nodes(aIndex1(0)).Nodes(aIndex1(1)).Nodes(aIndex1(2)).Nodes(aIndex1(3)).Nodes(aIndex1(4)).Nodes(aIndex1(5)).Nodes(aIndex1(6)).Nodes(aIndex1(7)).Nodes.Add(oNode1)
                            oNode1.ImageIndex = 3
                        Case 9
                            TV1.Nodes(aIndex1(0)).Nodes(aIndex1(1)).Nodes(aIndex1(2)).Nodes(aIndex1(3)).Nodes(aIndex1(4)).Nodes(aIndex1(5)).Nodes(aIndex1(6)).Nodes(aIndex1(7)).Nodes(aIndex1(8)).Nodes.Add(oNode1)
                            oNode1.ImageIndex = 0
                        Case 10
                            TV1.Nodes(aIndex1(0)).Nodes(aIndex1(1)).Nodes(aIndex1(2)).Nodes(aIndex1(3)).Nodes(aIndex1(4)).Nodes(aIndex1(5)).Nodes(aIndex1(6)).Nodes(aIndex1(7)).Nodes(aIndex1(8)).Nodes(aIndex1(9)).Nodes.Add(oNode1)
                            oNode1.ImageIndex = 3
                        Case 11
                            TV1.Nodes(aIndex1(0)).Nodes(aIndex1(1)).Nodes(aIndex1(2)).Nodes(aIndex1(3)).Nodes(aIndex1(4)).Nodes(aIndex1(5)).Nodes(aIndex1(6)).Nodes(aIndex1(7)).Nodes(aIndex1(8)).Nodes(aIndex1(9)).Nodes(aIndex1(10)).Nodes.Add(oNode1)
                            oNode1.ImageIndex = 3
                        Case 12
                            TV1.Nodes(aIndex1(0)).Nodes(aIndex1(1)).Nodes(aIndex1(2)).Nodes(aIndex1(3)).Nodes(aIndex1(4)).Nodes(aIndex1(5)).Nodes(aIndex1(6)).Nodes(aIndex1(7)).Nodes(aIndex1(8)).Nodes(aIndex1(9)).Nodes(aIndex1(10)).Nodes(aIndex1(11)).Nodes.Add(oNode1)
                            oNode1.ImageIndex = 3
                        Case 13
                            TV1.Nodes(aIndex1(0)).Nodes(aIndex1(1)).Nodes(aIndex1(2)).Nodes(aIndex1(3)).Nodes(aIndex1(4)).Nodes(aIndex1(5)).Nodes(aIndex1(6)).Nodes(aIndex1(7)).Nodes(aIndex1(8)).Nodes(aIndex1(9)).Nodes(aIndex1(10)).Nodes(aIndex1(11)).Nodes(aIndex1(12)).Nodes.Add(oNode1)
                            oNode1.ImageIndex = 3
                        Case 14
                            TV1.Nodes(aIndex1(0)).Nodes(aIndex1(1)).Nodes(aIndex1(2)).Nodes(aIndex1(3)).Nodes(aIndex1(4)).Nodes(aIndex1(5)).Nodes(aIndex1(6)).Nodes(aIndex1(7)).Nodes(aIndex1(8)).Nodes(aIndex1(9)).Nodes(aIndex1(10)).Nodes(aIndex1(11)).Nodes(aIndex1(12)).Nodes(aIndex1(13)).Nodes.Add(oNode1)
                            oNode1.ImageIndex = 3
                        Case 15
                            TV1.Nodes(aIndex1(0)).Nodes(aIndex1(1)).Nodes(aIndex1(2)).Nodes(aIndex1(3)).Nodes(aIndex1(4)).Nodes(aIndex1(5)).Nodes(aIndex1(6)).Nodes(aIndex1(7)).Nodes(aIndex1(8)).Nodes(aIndex1(9)).Nodes(aIndex1(10)).Nodes(aIndex1(11)).Nodes(aIndex1(12)).Nodes(aIndex1(13)).Nodes(aIndex1(14)).Nodes.Add(oNode1)
                            oNode1.ImageIndex = 3
                        Case 16
                            TV1.Nodes(aIndex1(0)).Nodes(aIndex1(1)).Nodes(aIndex1(2)).Nodes(aIndex1(3)).Nodes(aIndex1(4)).Nodes(aIndex1(5)).Nodes(aIndex1(6)).Nodes(aIndex1(7)).Nodes(aIndex1(8)).Nodes(aIndex1(9)).Nodes(aIndex1(10)).Nodes(aIndex1(11)).Nodes(aIndex1(12)).Nodes(aIndex1(13)).Nodes(aIndex1(14)).Nodes(aIndex1(15)).Nodes.Add(oNode1)
                            oNode1.ImageIndex = 3
                        Case 17
                            TV1.Nodes(aIndex1(0)).Nodes(aIndex1(1)).Nodes(aIndex1(2)).Nodes(aIndex1(3)).Nodes(aIndex1(4)).Nodes(aIndex1(5)).Nodes(aIndex1(6)).Nodes(aIndex1(7)).Nodes(aIndex1(8)).Nodes(aIndex1(9)).Nodes(aIndex1(10)).Nodes(aIndex1(11)).Nodes(aIndex1(12)).Nodes(aIndex1(13)).Nodes(aIndex1(14)).Nodes(aIndex1(15)).Nodes(aIndex1(16)).Nodes.Add(oNode1)
                            oNode1.ImageIndex = 3
                        Case 18
                            TV1.Nodes(aIndex1(0)).Nodes(aIndex1(1)).Nodes(aIndex1(2)).Nodes(aIndex1(3)).Nodes(aIndex1(4)).Nodes(aIndex1(5)).Nodes(aIndex1(6)).Nodes(aIndex1(7)).Nodes(aIndex1(8)).Nodes(aIndex1(9)).Nodes(aIndex1(10)).Nodes(aIndex1(11)).Nodes(aIndex1(12)).Nodes(aIndex1(13)).Nodes(aIndex1(14)).Nodes(aIndex1(15)).Nodes(aIndex1(16)).Nodes(aIndex1(17)).Nodes.Add(oNode1)
                            oNode1.ImageIndex = 3
                        Case 19
                            TV1.Nodes(aIndex1(0)).Nodes(aIndex1(1)).Nodes(aIndex1(2)).Nodes(aIndex1(3)).Nodes(aIndex1(4)).Nodes(aIndex1(5)).Nodes(aIndex1(6)).Nodes(aIndex1(7)).Nodes(aIndex1(8)).Nodes(aIndex1(9)).Nodes(aIndex1(10)).Nodes(aIndex1(11)).Nodes(aIndex1(12)).Nodes(aIndex1(13)).Nodes(aIndex1(14)).Nodes(aIndex1(15)).Nodes(aIndex1(16)).Nodes(aIndex1(17)).Nodes(aIndex1(18)).Nodes.Add(oNode1)
                            oNode1.ImageIndex = 3
                        Case 20
                            TV1.Nodes(aIndex1(0)).Nodes(aIndex1(1)).Nodes(aIndex1(2)).Nodes(aIndex1(3)).Nodes(aIndex1(4)).Nodes(aIndex1(5)).Nodes(aIndex1(6)).Nodes(aIndex1(7)).Nodes(aIndex1(8)).Nodes(aIndex1(9)).Nodes(aIndex1(10)).Nodes(aIndex1(11)).Nodes(aIndex1(12)).Nodes(aIndex1(13)).Nodes(aIndex1(14)).Nodes(aIndex1(15)).Nodes(aIndex1(16)).Nodes(aIndex1(17)).Nodes(aIndex1(18)).Nodes(aIndex1(19)).Nodes.Add(oNode1)
                            oNode1.ImageIndex = 3
                    End Select
                    'Hago la recursividad para llenar mis datos de la tabla en n niveles llamando la funcion
                    BUSCAR_ITEMS(dr("ID_NODO"))
                    B += 1
                Next
            End If
            nLevel1 -= 1
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim xn1 As TreeNode
    Private Sub RECORRER_Y_CARGAR_NODOS_1()
        Dim nodes1 As TreeNodeCollection = TV1.Nodes
        For Each n1 As TreeNode In nodes1
            xn1 = n1
            Call VERIFICAR_ESTADO_DE_NODO_EN_LA_TABLA_1()
            If ESTADO_DE_NODO_EN_LA_TABLA_1 = True Then
                n1.Checked = True
            Else
                n1.Checked = False
            End If
            RECORRER_Y_CARGAR_NODOS_SELECCIONADOS_1(n1)
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
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [MAESTRO DE NODOS USUARIOS] WHERE ID_USUARIO = " & uID_USUARIO & " AND ID_NODO = " & xn1.Tag & " AND ACCESO = '1'", Conexion.CxRRHH)
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
    Dim tn1 As TreeNode
    Private Sub RECORRER_Y_CARGAR_NODOS_SELECCIONADOS_1(treeNode1 As TreeNode)
        Try
            For Each tn1 In treeNode1.Nodes
                Call VERIFICAR_ESTADO_DE_NODO_EN_LA_TABLA()
                If ESTADO_DE_NODO_EN_LA_TABLA = True Then
                    tn1.Checked = True
                Else
                    tn1.Checked = False
                End If
                RECORRER_Y_CARGAR_NODOS_SELECCIONADOS_1(tn1)
            Next
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
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
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [MAESTRO DE NODOS USUARIOS] WHERE ID_USUARIO = " & uID_USUARIO & " AND ID_NODO = " & tn1.Tag & " AND ACCESO = '1'", Conexion.CxRRHH)
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
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ''Call VERIFICAR_ACCESOS("030") : If HAY_ACCESO = False Then : Exit Sub : End If
        For Each Nodo As TreeNode In TV1.Nodes
            SelectAllNodes(Nodo, True)
        Next
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ''Call VERIFICAR_ACCESOS("031") : If HAY_ACCESO = False Then : Exit Sub : End If
        For Each Nodo As TreeNode In TV1.Nodes
            SelectAllNodes(Nodo, False)
        Next
    End Sub
    Private Sub SelectAllNodes(Nodo As TreeNode, Checked As Boolean)
        Nodo.Checked = Checked
        For Each SubNodo As TreeNode In Nodo.Nodes
            SelectAllNodes(SubNodo, Checked)
        Next
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ''Call VERIFICAR_ACCESOS("032") : If HAY_ACCESO = False Then : Exit Sub : End If
        Me.Cursor = Cursors.WaitCursor
        If uID_USUARIO <> 0 Then
            Call RECORRER_NODOS()
        End If
        Me.Cursor = Cursors.Default
    End Sub
    Dim tn As TreeNode
    Private Sub RECORRER_NODOS()
        Dim nodes As TreeNodeCollection = TV1.Nodes
        For Each n As TreeNode In nodes
            RECORRER_NODOS_SELECCIONADOS(n)
        Next
    End Sub
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
    Private Sub ACTUALIZAR_TABLA_1()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE NODOS USUARIOS] SET ACCESO = '0' WHERE ID_NODO = " & tn.Tag & " and ID_USUARIO = " & uID_USUARIO & ""
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
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE NODOS USUARIOS] SET ACCESO = '1' WHERE ID_NODO = " & tn.Tag & " and ID_USUARIO = " & uID_USUARIO & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ''Call VERIFICAR_ACCESOS("029") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Accesos_Base.ShowDialog()
        Me.Text = "Accesos <" & uCUENTA & ">"
        Me.Cursor = Cursors.WaitCursor
        Call CARGAR_TV_1()
        TV1.Nodes(0).ForeColor = Color.DarkGreen : TV1.Nodes(0).Expand()
        TV1.Nodes(0).Nodes(0).ForeColor = Color.Blue : TV1.Nodes(0).Nodes(0).Expand()
        TV1.Nodes(0).Nodes(1).ForeColor = Color.Blue : TV1.Nodes(0).Nodes(1).Expand()
        TV1.Nodes(0).Nodes(2).ForeColor = Color.Blue : TV1.Nodes(0).Nodes(2).Expand()
        TV1.Nodes(0).Nodes(3).ForeColor = Color.Blue : TV1.Nodes(0).Nodes(3).Expand()
        TV1.Nodes(0).Nodes(4).ForeColor = Color.Blue : TV1.Nodes(0).Nodes(4).Expand()
        TV1.Nodes(0).Nodes(5).ForeColor = Color.Blue : TV1.Nodes(0).Nodes(5).Expand()
        TV1.Nodes(0).Nodes(5).Nodes(2).ForeColor = Color.Red
        TV1.SelectedNode = TV1.Nodes(0)
        Me.Cursor = Cursors.Default
        Call RECORRER_Y_CARGAR_NODOS_1()
        TV1.ExpandAll()
        TV1.SelectedNode = TV1.Nodes(0)
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ''Call VERIFICAR_ACCESOS("033") : If HAY_ACCESO = False Then : Exit Sub : End If
        Frm_Accesos_Asignar.ShowDialog()
    End Sub
End Class