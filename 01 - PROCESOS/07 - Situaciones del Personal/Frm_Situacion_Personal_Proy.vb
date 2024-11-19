Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports CrystalDecisions.[Shared].Json

Public Class Frm_Situacion_Personal_Proy


    ' Clase SituacionPersonal definida dentro del formulario
    Public Class SituacionPersonal
        Public Property ID_M_A As Integer
        Public Property ID_M_P As Integer
        Public Property DIA As Date
        Public Property ID_T_SIT_P As Integer
        Public Property ID_SIT_P As Integer
        Public Property VALOR_SIT As Decimal
        Public Property VALOR_DIA As Decimal
        Public Property OBSERVACIONES As String
        Public Property DETALLE_1 As Object
        Public Property DETALLE_2 As Object
        Public Property DETALLE_3 As Object
        Public Property USUARIO_ACT As String
        Public Property FECHA_ACT As DateTime
    End Class
    Private Sub Frm_Situacion_Personal_Proy_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Situacion_Personal_Proy_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Me.Close()
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW()
        ' Verifica y configura la conexión
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()

        ' Cargar datos en el DataSet
        Dim MiAdaptador As New SqlClient.SqlDataAdapter("select * from [VISTA MAESTRO DE PROYECCION] where ID_M_P = " & id_EMPLEADO & " ORDER BY FECHA_PROY", Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE PROYECCION]")

        ' Configurar el DataGridView
        DGV.Columns.Clear() ' Limpia el DataGridView
        DGV.AutoGenerateColumns = False ' Desactivar la generación automática de columnas

        ' Añadir columna de casilla de verificación en el índice 0
        Dim checkBoxColumn As New DataGridViewCheckBoxColumn()
        checkBoxColumn.HeaderText = "Seleccionar"
        checkBoxColumn.Width = 30
        checkBoxColumn.Name = "checkBoxColumn"
        DGV.Columns.Add(checkBoxColumn)

        ' Crear y añadir las columnas para los datos desde el índice 1
        For Each column As DataColumn In MiDataSet.Tables("[VISTA MAESTRO DE PROYECCION]").Columns
            Dim dgvColumn As New DataGridViewTextBoxColumn()
            dgvColumn.HeaderText = column.ColumnName
            dgvColumn.DataPropertyName = column.ColumnName
            DGV.Columns.Add(dgvColumn)
        Next

        ' Establecer el DataSource
        DGV.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE PROYECCION]")

        ' Cerrar la conexión si está abierta
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub

    Private Sub ARMAR_DATAGRIDVIEW()
        ' Configuración inicial del DataGridView
        Me.DGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.DGV.MultiSelect = True

        Me.DGV.Columns(0).HeaderText = "Seleccionar"
        Me.DGV.Columns(0).Width = 70
        Me.DGV.Columns(0).Visible = True
        Me.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(1).HeaderText = "Id"
        Me.DGV.Columns(1).Width = 30
        Me.DGV.Columns(1).Visible = True
        Me.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV.Columns(1).DefaultCellStyle.BackColor = Color.Black

        ' Configuración de las columnas (ajusta según tu necesidad)
        Me.DGV.Columns(2).HeaderText = "ID_M_P"
        Me.DGV.Columns(2).Width = 0
        Me.DGV.Columns(2).Visible = False

        Me.DGV.Columns(3).HeaderText = "ID_SIT_P"
        Me.DGV.Columns(3).Width = 0
        Me.DGV.Columns(3).Visible = False

        Me.DGV.Columns(4).HeaderText = "Situacion"
        Me.DGV.Columns(4).Width = 150
        Me.DGV.Columns(4).Visible = True
        Me.DGV.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DGV.Columns(4).DefaultCellStyle.BackColor = Color.LightBlue

        Me.DGV.Columns(5).HeaderText = "Valor Situacion"
        Me.DGV.Columns(5).Width = 100
        Me.DGV.Columns(5).Visible = True
        Me.DGV.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DGV.Columns(6).HeaderText = "Fecha Proyeccion"
        Me.DGV.Columns(6).Width = 70
        Me.DGV.Columns(6).Visible = True
        Me.DGV.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(6).DefaultCellStyle.BackColor = Color.Khaki

        Me.DGV.Columns(7).HeaderText = "Observaciones"
        Me.DGV.Columns(7).Width = 370
        Me.DGV.Columns(7).Visible = True
        Me.DGV.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

    End Sub
    'Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW()
    '    Dim I As Integer
    '    For I = 0 To Me.DGV.RowCount - 1
    '        If Me.DGV.Item(0, I).Value = proyID_PROY Then
    '            Me.DGV.Rows(I).Selected = True
    '            Me.DGV.CurrentCell = Me.DGV.Item(0, I)
    '            Exit Sub
    '        End If
    '    Next
    'End Sub
    Dim MENSAJE_1
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click 'Botón Eliminar'
        ' Recopilar IDs seleccionados
        Dim idsSeleccionados As New List(Of Integer)
        For Each row As DataGridViewRow In DGV.Rows
            Dim isSelected As Boolean = Convert.ToBoolean(row.Cells("checkBoxColumn").Value)
            If isSelected Then
                idsSeleccionados.Add(Convert.ToInt32(row.Cells(1).Value)) ' Asumiendo que la columna ID es la segunda
            End If
        Next

        ' Verifica si hay registros seleccionados
        If idsSeleccionados.Count = 0 Then
            MsgBox("Debe seleccionar al menos un registro para eliminar.", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If

        ' Confirmación para eliminar múltiples registros
        Dim MENSAJE = MsgBox("¿Está seguro de eliminar los registros seleccionados?", vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            Call ELIMINAR_REGISTROS(idsSeleccionados)
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
        End If
    End Sub
    'Private Sub ELIMINAR_REGISTROS(ids As List(Of Integer))
    '    If Conexion.CxRRHH.State = ConnectionState.Open Then
    '        Conexion.CBD_RRHH()
    '    End If
    '    Conexion.ABD_RRHH()

    '    Try
    '        ' Construye la consulta de eliminación
    '        Dim idsParaEliminar As String = String.Join(",", ids)
    '        Dim query As String = "DELETE FROM [MAESTRO DE PROYECCION] WHERE ID_PROY IN (" & idsParaEliminar & ")"
    '        Dim MiSqlCommand As New SqlCommand(query, Conexion.CxRRHH)
    '        MiSqlCommand.ExecuteNonQuery()
    '    Catch ex As Exception
    '        MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
    '    Finally
    '        If Conexion.CxRRHH.State = ConnectionState.Open Then
    '            Conexion.CBD_RRHH()
    '        End If
    '    End Try
    'End Sub

    'Private Sub DGV_Click(sender As Object, e As EventArgs) Handles DGV.Click
    '    If Me.DGV.RowCount > 0 Then
    '        For I = 0 To Me.DGV.RowCount - 1
    '            If DGV.Rows(I).Selected = True Then
    '                proyID_PROY = Me.DGV.Item(0, I).Value
    '                proyID_M_P = Me.DGV.Item(1, I).Value
    '                proyID_SIT_P = Me.DGV.Item(2, I).Value
    '                proyD_SIT_P = Me.DGV.Item(3, I).Value
    '                proyFECHA_PROY = Me.DGV.Item(4, I).Value
    '                proyOBSERVACIONES = Me.DGV.Item(5, I).Value
    '                Exit Sub
    '            End If
    '        Next
    '    End If
    'End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Frm_Situacion_Personal_Proy_Add.ShowDialog()
        'xxx_Frm_Situacion_Personal_Proyectar.ShowDialog()
        If CERRAR = False Then
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            'Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
            proyID_PROY = 0
        End If
    End Sub

    Private Sub ButtonSaveProyectionSituationPersonal_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' Abre la conexión si está cerrada
        If Conexion.CxRRHH.State <> ConnectionState.Open Then
            Conexion.ABD_RRHH()
        End If

        Try
            Dim fechaActual As DateTime = DateTime.Now
            Dim dataTable As New DataTable
            dataTable.Columns.Add("ID_M_A", GetType(Integer))
            dataTable.Columns.Add("ID_M_P", GetType(Integer))
            dataTable.Columns.Add("DIA", GetType(Date))
            dataTable.Columns.Add("ID_T_SIT_P", GetType(Integer))
            dataTable.Columns.Add("ID_SIT_P", GetType(Integer))
            dataTable.Columns.Add("VALOR_SIT", GetType(Decimal))
            dataTable.Columns.Add("VALOR_DIA", GetType(Decimal))
            dataTable.Columns.Add("OBSERVACIONES", GetType(String))
            dataTable.Columns.Add("DETALLE_1", GetType(String))
            dataTable.Columns.Add("DETALLE_2", GetType(String))
            dataTable.Columns.Add("DETALLE_3", GetType(String))
            dataTable.Columns.Add("USUARIO_ACT", GetType(String))
            dataTable.Columns.Add("FECHA_ACT", GetType(Date))

            ' Llena el DataTable con los datos del DataGridView
            For Each row As DataGridViewRow In Me.DGV.Rows
                If row.IsNewRow Then Continue For

                Dim situacion As New SituacionPersonal With {
                    .ID_M_A = 0,
                    .ID_M_P = Convert.ToInt32(row.Cells(2).Value),
                    .DIA = row.Cells(6).Value,
                    .ID_T_SIT_P = 2,
                    .ID_SIT_P = Convert.ToInt32(row.Cells(3).Value),
                    .VALOR_SIT = Convert.ToDecimal(row.Cells(5).Value),
                    .VALOR_DIA = 0.0833333333333333D,
                    .OBSERVACIONES = row.Cells(7).Value.ToString(),
                    .DETALLE_1 = DBNull.Value,
                    .DETALLE_2 = DBNull.Value,
                    .DETALLE_3 = DBNull.Value,
                    .USUARIO_ACT = isuUSUARIO,
                    .FECHA_ACT = Convert.ToDateTime(fechaActual)
                }

                dataTable.Rows.Add(
                    situacion.ID_M_A,
                    situacion.ID_M_P,
                    situacion.DIA,
                    situacion.ID_T_SIT_P,
                    situacion.ID_SIT_P,
                    situacion.VALOR_SIT,
                    situacion.VALOR_DIA,
                    situacion.OBSERVACIONES,
                    situacion.DETALLE_1,
                    situacion.DETALLE_2,
                    situacion.DETALLE_3,
                    situacion.USUARIO_ACT,
                    situacion.FECHA_ACT
                )
            Next

            ' Obtén los valores únicos de ID_M_P y formatea para usar en una consulta IN
            Dim ID_M_P As Integer = Convert.ToInt32(String.Join(",", dataTable.AsEnumerable().Select(Function(dr) dr.Field(Of Integer)("ID_M_P")).Distinct()))

            ' Obtén las fechas en formato 'yyyy-MM-dd' y rodeadas de comillas simples para la consulta
            Dim fechas As String = String.Join(",", dataTable.AsEnumerable().Select(Function(dr) "'" & dr.Field(Of DateTime)("DIA").ToString("yyyy-MM-dd") & "'").Distinct())

            ' Construye la consulta utilizando IN para ID_M_P y fechas
            Dim queryExistentes As String = "SELECT DIA FROM [dbo].[MAESTRO DE SITUACION DE PERSONAL] WHERE DIA IN (" & fechas & ") AND ID_M_P IN (" & ID_M_P & ")"

            Dim fechasExistentes As New HashSet(Of DateTime)
            Using cmdExistentes As New SqlCommand(queryExistentes, Conexion.CxRRHH)
                Using reader As SqlDataReader = cmdExistentes.ExecuteReader()
                    While reader.Read()
                        fechasExistentes.Add(reader.GetDateTime(0))
                    End While
                End Using
            End Using

            ' Paso 2: Dividir los datos en dos tablas: para insertar y para actualizar
            Dim dataTableInsertar As DataTable = dataTable.Clone()
            Dim dataTableActualizar As DataTable = dataTable.Clone()

            For Each row As DataRow In dataTable.Rows
                If fechasExistentes.Contains(row.Field(Of DateTime)("DIA")) Then
                    dataTableActualizar.ImportRow(row)
                Else
                    dataTableInsertar.ImportRow(row)
                End If
            Next

            ' Inicia una transacción SQL
            Using transaction As SqlTransaction = Conexion.CxRRHH.BeginTransaction()
                Try
                    ' Paso 3: Insertar nuevos registros con SqlBulkCopy
                    If dataTableInsertar.Rows.Count > 0 Then
                        Using bulkCopy As New SqlBulkCopy(Conexion.CxRRHH, SqlBulkCopyOptions.Default, transaction)
                            bulkCopy.DestinationTableName = "[dbo].[MAESTRO DE SITUACION DE PERSONAL]"
                            bulkCopy.BatchSize = 1000
                            bulkCopy.ColumnMappings.Add("ID_M_A", "ID_M_A")
                            bulkCopy.ColumnMappings.Add("ID_M_P", "ID_M_P")
                            bulkCopy.ColumnMappings.Add("DIA", "DIA")
                            bulkCopy.ColumnMappings.Add("ID_T_SIT_P", "ID_T_SIT_P")
                            bulkCopy.ColumnMappings.Add("ID_SIT_P", "ID_SIT_P")
                            bulkCopy.ColumnMappings.Add("VALOR_SIT", "VALOR_SIT")
                            bulkCopy.ColumnMappings.Add("VALOR_DIA", "VALOR_DIA")
                            bulkCopy.ColumnMappings.Add("OBSERVACIONES", "OBSERVACIONES")
                            bulkCopy.ColumnMappings.Add("DETALLE_1", "DETALLE_1")
                            bulkCopy.ColumnMappings.Add("DETALLE_2", "DETALLE_2")
                            bulkCopy.ColumnMappings.Add("DETALLE_3", "DETALLE_3")
                            bulkCopy.ColumnMappings.Add("USUARIO_ACT", "USUARIO_ACT")
                            bulkCopy.ColumnMappings.Add("FECHA_ACT", "FECHA_ACT")
                            bulkCopy.WriteToServer(dataTableInsertar)
                        End Using
                    End If

                    ' Paso 4: Actualizar registros existentes
                    For Each row As DataRow In dataTableActualizar.Rows
                        Dim updateQuery As String = "UPDATE [dbo].[MAESTRO DE SITUACION DE PERSONAL] SET " &
                            "ID_M_P = @ID_M_P, ID_T_SIT_P = @ID_T_SIT_P, ID_SIT_P = @ID_SIT_P, " &
                            "VALOR_SIT = @VALOR_SIT, VALOR_DIA = @VALOR_DIA, OBSERVACIONES = @OBSERVACIONES, " &
                            "DETALLE_1 = @DETALLE_1, DETALLE_2 = @DETALLE_2, DETALLE_3 = @DETALLE_3, " &
                            "USUARIO_ACT = @USUARIO_ACT, FECHA_ACT = @FECHA_ACT " &
                            "WHERE DIA = @DIA"

                        Using cmdUpdate As New SqlCommand(updateQuery, Conexion.CxRRHH, transaction)
                            cmdUpdate.Parameters.AddWithValue("@ID_M_P", row("ID_M_P"))
                            cmdUpdate.Parameters.AddWithValue("@ID_T_SIT_P", row("ID_T_SIT_P"))
                            cmdUpdate.Parameters.AddWithValue("@ID_SIT_P", row("ID_SIT_P"))
                            cmdUpdate.Parameters.AddWithValue("@VALOR_SIT", row("VALOR_SIT"))
                            cmdUpdate.Parameters.AddWithValue("@VALOR_DIA", row("VALOR_DIA"))
                            cmdUpdate.Parameters.AddWithValue("@OBSERVACIONES", row("OBSERVACIONES"))
                            cmdUpdate.Parameters.AddWithValue("@DETALLE_1", row("DETALLE_1"))
                            cmdUpdate.Parameters.AddWithValue("@DETALLE_2", row("DETALLE_2"))
                            cmdUpdate.Parameters.AddWithValue("@DETALLE_3", row("DETALLE_3"))
                            cmdUpdate.Parameters.AddWithValue("@USUARIO_ACT", row("USUARIO_ACT"))
                            cmdUpdate.Parameters.AddWithValue("@FECHA_ACT", row("FECHA_ACT"))
                            cmdUpdate.Parameters.AddWithValue("@DIA", row("DIA"))
                            cmdUpdate.ExecuteNonQuery()
                        End Using
                    Next

                    ' Confirma la transacción si todo es exitoso
                    transaction.Commit()
                    MsgBox("Datos insertados y actualizados exitosamente.", vbInformation, "Operación Exitosa")
                    For Each row As DataGridViewRow In DGV.Rows
                        ' Omite la fila de "nuevo" en el DataGridView
                        If row.IsNewRow Then Continue For

                        ' Cambia el valor de la celda en la columna 0 (debería ser un valor booleano) a True
                        row.Cells(0).Value = True
                    Next
                    EliminarRegistrosMarcados()
                Catch ex As Exception
                    ' En caso de error, revierte la transacción
                    transaction.Rollback()
                    MsgBox("Error al insertar datos: " & ex.Message, vbExclamation, "Error")
                End Try
            End Using

        Catch ex As Exception
            MsgBox("Error de conexión: " & ex.Message, vbExclamation, "Error")
        Finally
            ' Cierra la conexión si está abierta
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub

    ' Método para eliminar los registros de las filas cuyo valor en la columna 0 es True
    Private Sub EliminarRegistrosMarcados()
        Dim idsParaEliminar As New List(Of Integer)()

        For Each row As DataGridViewRow In DGV.Rows
            If row.IsNewRow Then Continue For

            If Convert.ToBoolean(row.Cells(0).Value) = True Then
                If Not IsDBNull(row.Cells(1).Value) Then
                    idsParaEliminar.Add(Convert.ToInt32(row.Cells(1).Value))
                End If
            End If
        Next
        If idsParaEliminar.Count > 0 Then
            ELIMINAR_REGISTROS(idsParaEliminar)
        Else
            MsgBox("No hay registros marcados para eliminar.", vbInformation, "Mensaje del Sistema")
        End If
    End Sub

    ' Método de eliminación de registros en la base de datos con transacción
    Private Sub ELIMINAR_REGISTROS(ids As List(Of Integer))
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()

        ' Iniciar transacción
        Dim transaction As SqlTransaction = Nothing
        Try
            ' Abrir la conexión si no está abierta
            If Conexion.CxRRHH.State = ConnectionState.Closed Then
                Conexion.CxRRHH.Open()
            End If

            ' Iniciar la transacción
            transaction = Conexion.CxRRHH.BeginTransaction()

            ' Construye la consulta de eliminación
            Dim idsParaEliminar As String = String.Join(",", ids)
            Dim query As String = "DELETE FROM [MAESTRO DE PROYECCION] WHERE ID_PROY IN (" & idsParaEliminar & ")"
            Dim MiSqlCommand As New SqlCommand(query, Conexion.CxRRHH)
            MiSqlCommand.Transaction = transaction ' Asocia la transacción al comando

            ' Ejecuta la consulta
            MiSqlCommand.ExecuteNonQuery()

            ' Si todo fue exitoso, confirma la transacción
            transaction.Commit()
            MsgBox("Registros eliminados con éxito.", vbInformation, "Mensaje del Sistema")

        Catch ex As Exception
            ' Si algo falla, revierte la transacción
            If transaction IsNot Nothing Then
                transaction.Rollback()
            End If
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            ' Asegúrate de cerrar la conexión si no está cerrada
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CxRRHH.Close()
            End If
        End Try
    End Sub

    Private Sub DGV_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV.CellContentClick
        ' Verifica si se hizo clic en la columna de checkbox
        If e.ColumnIndex = 0 AndAlso e.RowIndex >= 0 Then
            ' Cambia el valor del checkbox al hacer clic
            Dim checkCell As DataGridViewCheckBoxCell = CType(DGV.Rows(e.RowIndex).Cells(0), DataGridViewCheckBoxCell)
            checkCell.Value = Not CBool(checkCell.Value) ' Cambia el valor actual
        End If
    End Sub
End Class