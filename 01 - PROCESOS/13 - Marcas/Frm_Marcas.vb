Imports System.Data.SqlClient
Imports System.Reflection

Public Class Frm_Marcas

    ' Diccionario global para almacenar los valores seleccionados por cada ComboBox
    Private seleccionDatos As New Dictionary(Of String, Dictionary(Of String, String))
    Private ultimoComboBoxSeleccionado

    Private Sub Frm_Marcas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Definir parámetros de la consulta (vacío en este caso)

        ' Definir parámetros de la consulta
        '    Dim parametros As New Dictionary(Of String, Object) From {
        '    {"@edad", 25},
        '    {"@ciudad", "Madrid"}
        '}

        '    ' Definir los INNER JOIN o LEFT JOIN
        '    Dim joins As New List(Of String) From {
        '    "INNER JOIN Direcciones d ON c.id = d.id_cliente",
        '    "LEFT JOIN Compras cmp ON c.id = cmp.id_cliente"
        '}

        ' Obtener la fecha actual
        Dim fechaActual As DateTime = DateTime.Now

        ' Establecer DateTimePicker2 con la fecha actual
        DateTimePicker2.Value = fechaActual

        ' Establecer DateTimePicker1 con el primer día del mes actual
        DateTimePicker1.Value = New DateTime(fechaActual.Year, fechaActual.Month, 1)


        seleccionDatos = New Dictionary(Of String, Dictionary(Of String, String)) From {
    {"n1.SIGLAS", New Dictionary(Of String, String) From {{"ComboBox1", "NULL"}}},
    {"n2.SIGLAS", New Dictionary(Of String, String) From {{"ComboBox2", "NULL"}}},
    {"n3.SIGLAS", New Dictionary(Of String, String) From {{"ComboBox3", "NULL"}}},
    {"n4.SIGLAS", New Dictionary(Of String, String) From {{"ComboBox4", "NULL"}}},
    {"n5.SIGLAS", New Dictionary(Of String, String) From {{"ComboBox5", "NULL"}}},
    {"n6.SIGLAS", New Dictionary(Of String, String) From {{"ComboBox6", "NULL"}}},
    {"n7.SIGLAS", New Dictionary(Of String, String) From {{"ComboBox7", "NULL"}}},
    {"n8.SIGLAS", New Dictionary(Of String, String) From {{"ComboBox8", "NULL"}}}
}

        Dim parametros As New Dictionary(Of String, Object) From {}
        Dim joins As New List(Of String) From {}

        ' Ejecutar consulta para obtener datos
        Dim dt As DataTable = ConsultarDatos("dbo.[CAT DE ESTRUCTURA NIVEL 1]", joins, "*", "", parametros, "", "ID_N1 ASC")

        ' Rellenar el ComboBox1 con los resultados
        ConfigurarComboBoxPorDefecto()
        ConfigurarDataGridView()
        ' Llenar ComboBox1 con codigo_interno como identificador y DESCRIPCION como visual
        LlenarComboBox(ComboBox1, dt, "DESCRIPCION", "SIGLAS", seleccionarPrimero:=True)
    End Sub

    ''' <summary>
    ''' Configura los ComboBox y CheckBox del formulario con valores por defecto.
    ''' </summary>
    Private Sub ConfigurarComboBoxPorDefecto()
        ' Lista de ComboBox y CheckBox a configurar (1-8)
        Dim comboBoxes As ComboBox() = {ComboBox1, ComboBox2, ComboBox3, ComboBox4, ComboBox5, ComboBox6, ComboBox7, ComboBox8}
        Dim checkBoxes As CheckBox() = {CheckBox1, CheckBox2, CheckBox3, CheckBox4, CheckBox5, CheckBox6, CheckBox7, CheckBox8}

        For i As Integer = 0 To comboBoxes.Length - 1
            ' Configurar ComboBox
            Dim cb As ComboBox = comboBoxes(i)
            cb.DropDownStyle = ComboBoxStyle.DropDownList ' Hacer que no sean editables
            cb.Items.Clear()
            cb.Items.Add("SELECCIONE") ' Agregar opción por defecto
            cb.SelectedIndex = 0 ' Seleccionar "SELECCIONE" por defecto

            ' Configurar CheckBox
            Dim ch As CheckBox = checkBoxes(i)
            ch.Enabled = False ' Desactivar CheckBox por defecto

            ' Agregar evento de cambio de estado del CheckBox
            AddHandler ch.CheckedChanged, AddressOf CheckBox_CheckedChanged

            ' Agregar evento de selección en ComboBox
            AddHandler cb.SelectedIndexChanged, AddressOf ComboBox_SelectedIndexChanged
        Next
    End Sub

    Private Sub ConfigurarDataGridView()
        With DataGridView1
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells ' Ajusta todas las columnas al contenido
            .AutoResizeColumns() ' Forzar ajuste automático
        End With
    End Sub

    ''' <summary>
    ''' Valida que los DateTimePickers tengan fechas correctas y en orden lógico.
    ''' </summary>
    ''' <returns>True si las fechas son válidas, False si hay algún problema.</returns>
    Private Function ValidarFechas() As Boolean
        ' Verificar que los DateTimePickers contienen fechas válidas
        If Not DateTimePicker1.Checked OrElse Not DateTimePicker2.Checked Then
            MessageBox.Show("Debe seleccionar fechas válidas en ambos campos.", "Validación de Fechas", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If

        ' Obtener valores de los DateTimePickers
        Dim fechaInicio As DateTime = DateTimePicker1.Value
        Dim fechaFin As DateTime = DateTimePicker2.Value

        ' Validar que FechaInicio no sea mayor que FechaFin
        If fechaInicio > fechaFin Then
            MessageBox.Show("La fecha de inicio no puede ser mayor que la fecha de fin. Ajuste las fechas.", "Error de Fechas", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If

        ' Si todas las validaciones pasan, devolver True
        Return True
    End Function


    ''' <summary>
    ''' Método para llenar un ComboBox con datos (codigo_interno, DESCRIPCION).
    ''' </summary>
    ''' <param name="cb">ComboBox a llenar</param>
    ''' <param name="dt">DataTable con los datos</param>
    ''' <param name="campoTexto">Campo a mostrar en el ComboBox</param>
    ''' <param name="campoValor">Campo identificador interno</param>
    ''' <param name="seleccionarPrimero">Si True, selecciona automáticamente el primer valor</param>
    Private Sub LlenarComboBox(cb As ComboBox, dt As DataTable, campoTexto As String, campoValor As String, Optional seleccionarPrimero As Boolean = False)
        Try
            cb.Items.Clear()
            cb.DataSource = Nothing
            cb.DisplayMember = campoTexto
            cb.ValueMember = campoValor

            Dim dtClon As DataTable = dt.Clone()

            Dim newRow As DataRow = dtClon.NewRow()

            If dtClon.Columns(campoValor).DataType Is GetType(Double) OrElse dtClon.Columns(campoValor).DataType Is GetType(Integer) Then
                newRow(campoValor) = DBNull.Value
            Else
                newRow(campoValor) = "0"
            End If

            newRow(campoTexto) = "SELECCIONE"
            dtClon.Rows.Add(newRow)

            dtClon.Merge(dt)

            cb.DataSource = dtClon

            If seleccionarPrimero AndAlso cb.Items.Count > 1 Then
                cb.SelectedIndex = 1
            Else
                cb.SelectedIndex = 0
            End If

        Catch ex As Exception
            MessageBox.Show("Error al llenar el ComboBox: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim comboBoxes As ComboBox() = {ComboBox1, ComboBox2, ComboBox3, ComboBox4, ComboBox5, ComboBox6, ComboBox7, ComboBox8}
        Dim checkBoxes As CheckBox() = {CheckBox1, CheckBox2, CheckBox3, CheckBox4, CheckBox5, CheckBox6, CheckBox7, CheckBox8}
        If Not ValidarFechas() Then
            Exit Sub ' Detener ejecución si las fechas no son válidas
        End If
        ' Identificar el ComboBox que activó el evento
        For i As Integer = 0 To comboBoxes.Length - 1
            If sender Is comboBoxes(i) Then
                ' Habilitar y marcar automáticamente el CheckBox correspondiente si la selección es válida
                If comboBoxes(i).SelectedIndex > 0 Then
                    checkBoxes(i).Enabled = True
                    'checkBoxes(i).Checked = True ' 🔹 Forzar activación del CheckBox
                Else
                    checkBoxes(i).Enabled = False
                    checkBoxes(i).Checked = False ' 🔹 Desmarcar si vuelve a SELECCIONE o vacío
                End If

                ' Limpiar todos los ComboBox siguientes para evitar datos erróneos
                For j As Integer = i + 1 To comboBoxes.Length - 1
                    comboBoxes(j).DataSource = Nothing
                    comboBoxes(j).Items.Clear()
                    comboBoxes(j).Enabled = False
                Next

                ' Verificar que haya un siguiente ComboBox para llenar
                If i < comboBoxes.Length - 1 AndAlso comboBoxes(i).SelectedIndex > 0 Then
                    Dim nivelActual As Integer = i + 1
                    Dim nivelSiguiente As Integer = i + 2
                    Dim tabla As String = $"dbo.[CAT DE ESTRUCTURA NIVEL {nivelSiguiente}]"
                    Dim orderBy As String = $"ID_N{nivelSiguiente} ASC"

                    Dim row As DataRowView = DirectCast(comboBoxes(i).SelectedItem, DataRowView)
                    Dim valor = row($"ID_N{nivelActual}").ToString()

                    ' Construcción de la condición WHERE
                    Dim whereCondition As String = $"ID_N{nivelActual} = @param{nivelActual}"
                    Dim parametros As New Dictionary(Of String, Object) From {
                    {$"@param{nivelActual}", valor}
                }

                    ' Ejecutar la consulta solo para el siguiente nivel
                    Dim dt As DataTable = ConsultarDatos(tabla, New List(Of String), "*", whereCondition, parametros, "", orderBy)
                    checkBoxes(i).Checked = False
                    ' Solo llenar el siguiente ComboBox si hay datos disponibles
                    If dt.Rows.Count > 0 Then
                        LlenarComboBox(comboBoxes(i + 1), dt, "DESCRIPCION", $"ID_N{nivelSiguiente}", False)
                        comboBoxes(i + 1).Enabled = True
                    Else
                        comboBoxes(i + 1).Enabled = False
                    End If
                End If

                ' Romper el bucle después de encontrar el ComboBox activado
                Exit For
            End If
        Next
        ultimoComboBoxSeleccionado = DirectCast(sender, ComboBox)
    End Sub


    Private Sub CheckBox_CheckedChanged(sender As Object, e As EventArgs)
        Call VERIFICAR_ACCESOS("376") : If HAY_ACCESO = False Then : Exit Sub : End If
        Dim comboBoxes As ComboBox() = {ComboBox1, ComboBox2, ComboBox3, ComboBox4, ComboBox5, ComboBox6, ComboBox7, ComboBox8}
        Dim checkBoxes As CheckBox() = {CheckBox1, CheckBox2, CheckBox3, CheckBox4, CheckBox5, CheckBox6, CheckBox7, CheckBox8}
        Dim keyMapping As String() = {"n1.SIGLAS", "n2.SIGLAS", "n3.SIGLAS", "n4.SIGLAS", "n5.SIGLAS", "n6.SIGLAS", "n7.SIGLAS", "n8.SIGLAS"} ' Llaves dinámicas

        ' Buscar el CheckBox que activó el evento
        For i As Integer = 0 To checkBoxes.Length - 1
            If sender Is checkBoxes(i) Then
                Dim valorIDN As String = keyMapping(i) ' Clave de la estructura en el diccionario

                ' Si el CheckBox está marcado (Checked = True)
                If checkBoxes(i).Checked Then
                    ' Obtener valores del ComboBox asociado
                    Dim valorTexto As String = ""
                    Dim valorInterno As String = ""

                    ' Asegurar que hay un ítem seleccionado y que no sea "SELECCIONE"
                    If comboBoxes(i).SelectedItem IsNot Nothing AndAlso comboBoxes(i).SelectedIndex > 0 Then
                        Dim row As DataRowView = DirectCast(comboBoxes(i).SelectedItem, DataRowView)
                        valorTexto = row("DESCRIPCION").ToString()
                        valorInterno = row("SIGLAS").ToString()

                        If valorInterno Is DBNull.Value Then
                            valorInterno = "NULL"
                        End If
                    Else
                        valorTexto = "SELECCIONE"
                        valorInterno = "NULL"
                    End If

                    ' Guardar los valores en el diccionario
                    If Not seleccionDatos.ContainsKey(valorIDN) Then
                        seleccionDatos(valorIDN) = New Dictionary(Of String, String)
                    End If

                    ' Guardar campo relacionado con el ComboBox
                    seleccionDatos(valorIDN)("ComboBox" & (i + 1)) = valorInterno

                Else ' Si el CheckBox es desmarcado (Checked = False)
                    If seleccionDatos.ContainsKey(valorIDN) AndAlso seleccionDatos(valorIDN).ContainsKey("ComboBox" & (i + 1)) Then
                        seleccionDatos(valorIDN)("ComboBox" & (i + 1)) = "NULL" ' Actualizar a NULL
                    End If
                End If

                ' Construir el WHERE y los parámetros después de actualizar la selección
                Dim whereCondition As String = ConstruirWhereCondition()
                Dim parametros As Dictionary(Of String, Object) = ConstruirParametros()

                ' Llamar a la consulta y actualizar los resultados en DataGridView1
                ActualizarDataGridView(whereCondition, parametros, seleccionDatos(valorIDN)("ComboBox" & (i + 1)))
            End If
        Next
    End Sub


    'Private Sub CheckBox_CheckedChanged(sender As Object, e As EventArgs)
    '    Dim comboBoxes As ComboBox() = {ComboBox1, ComboBox2, ComboBox3, ComboBox4, ComboBox5, ComboBox6, ComboBox7, ComboBox8}
    '    Dim checkBoxes As CheckBox() = {CheckBox1, CheckBox2, CheckBox3, CheckBox4, CheckBox5, CheckBox6, CheckBox7, CheckBox8}
    '    Dim keyMapping As String() = {"n1.SIGLAS", "n2.SIGLAS", "n3.SIGLAS", "n4.SIGLAS", "n5.SIGLAS", "n6.SIGLAS", "n7.SIGLAS", "n8.SIGLAS"} ' Llaves dinámicas

    '    ' Buscar el CheckBox que activó el evento
    '    For i As Integer = 0 To checkBoxes.Length - 1
    '        If sender Is checkBoxes(i) AndAlso checkBoxes(i).Checked Then
    '            ' Obtener valores del ComboBox asociado
    '            Dim valorTexto As String = ""
    '            Dim valorInterno As String = ""
    '            Dim valorIDN As String = keyMapping(i) ' ID_N asociado a este ComboBox

    '            ' Asegurar que hay un ítem seleccionado y que no sea "SELECCIONE"
    '            If comboBoxes(i).SelectedItem IsNot Nothing AndAlso comboBoxes(i).SelectedIndex > 0 Then
    '                Dim row As DataRowView = DirectCast(comboBoxes(i).SelectedItem, DataRowView)
    '                valorTexto = row("DESCRIPCION").ToString()
    '                valorInterno = row("SIGLAS").ToString()

    '                If valorInterno Is DBNull.Value Then
    '                    valorInterno = "NULL"
    '                End If
    '            Else
    '                valorTexto = "SELECCIONE"
    '                valorInterno = "NULL"
    '            End If

    '            ' Guardar los valores en el diccionario
    '            If Not seleccionDatos.ContainsKey(valorIDN) Then
    '                seleccionDatos(valorIDN) = New Dictionary(Of String, String)
    '            End If

    '            ' Guardar campo relacionado con el ComboBox
    '            seleccionDatos(valorIDN)("ComboBox" & (i + 1)) = valorInterno

    '            ' Construir el WHERE y los parámetros después de actualizar la selección
    '            Dim whereCondition As String = ConstruirWhereCondition()
    '            Dim parametros As Dictionary(Of String, Object) = ConstruirParametros()


    '            ' Llamar a la consulta y mostrar los resultados en DataGridView1
    '            ActualizarDataGridView(whereCondition, parametros, valorInterno)
    '        End If
    '    Next
    'End Sub



    ''' <summary>
    ''' Construye los parámetros para la consulta SQL.
    ''' </summary>
    Private Function ConstruirParametros() As Dictionary(Of String, Object)
        Dim parametros As New Dictionary(Of String, Object)

        For Each key As String In seleccionDatos.Keys
            Dim valor As String = seleccionDatos(key).Values.First()

            ' Solo agregar si el valor no es NULL
            If valor <> "NULL" Then
                ' Limpiar la clave eliminando '.' y '_'
                Dim keyLimpia As String = key.Replace(".", "").Replace("_", "")

                parametros.Add("@" & keyLimpia, valor)
            End If
        Next

        Return parametros
    End Function

    ''' <summary>
    ''' Construye la condición WHERE con los valores seleccionados.
    ''' </summary>
    Private Function ConstruirWhereCondition() As String
        Dim condiciones As New List(Of String)

        For Each key As String In seleccionDatos.Keys
            Dim valor As String = seleccionDatos(key).Values.First()

            ' Solo agregar al WHERE si el valor no es NULL
            If valor <> "NULL" Then
                ' Limpiar la clave eliminando '.' y '_'
                Dim keyLimpia As String = key.Replace(".", "").Replace("_", "")
                condiciones.Add($"{key} = @{keyLimpia}")
            End If
        Next

        ' Unir condiciones con AND
        Return If(condiciones.Count > 0, String.Join(" AND ", condiciones), "1=1") ' "1=1" evita errores si no hay condiciones
    End Function

    Private Function ActualizarDataGridView(whereCondition As String, parametros As Dictionary(Of String, Object), valor As String) As DataTable
        ' Definir los INNER JOIN o LEFT JOIN
        Dim joins As New List(Of String) From {
       "INNER JOIN dbo.[CAT DE TIPO DE ESTRUCTURA] tes ON mc.ID_T_ES = tes.ID_T_ES",
       "INNER JOIN dbo.[CAT DE CARGOS DE ESTRUCTURA] ces ON mc.ID_CARGO_ES = ces.ID_CARGO_ES",
       "INNER JOIN dbo.[CAT DE GRADO PLANTILLA] gp ON mc.ID_GP = gp.ID_GP",
       "INNER JOIN dbo.[CAT DE GRADO REAL] gr ON mc.ID_GR = gr.ID_GR",
       "INNER JOIN dbo.[CAT DE CATEGORIA SALARIAL] cs ON mc.ID_CAT_SALARIAL = cs.ID_CAT_SALARIAL",
       "INNER JOIN dbo.[MAESTRO DE PERSONAS] mp ON mc.ID_M_P = mp.ID_M_P",
       "INNER JOIN dbo.[CAT DE ESTABLECIMIENTO] est ON mc.ID_ESTABLECIMIENTO = est.ID_ESTABLECIMIENTO",
       "INNER JOIN dbo.[CAT DE ADMINISTRADORES] adm ON mp.ID_ADMIN = adm.ID_ADMIN",
       "LEFT JOIN dbo.[CAT DE ESTRUCTURA NIVEL 1] n1 ON mc.ID_N1 = n1.ID_N1",
       "LEFT JOIN dbo.[CAT DE ESTRUCTURA NIVEL 2] n2 ON mc.ID_N2 = n2.ID_N2",
       "LEFT JOIN dbo.[CAT DE ESTRUCTURA NIVEL 3] n3 ON mc.ID_N3 = n3.ID_N3",
       "LEFT JOIN dbo.[CAT DE ESTRUCTURA NIVEL 4] n4 ON mc.ID_N4 = n4.ID_N4",
       "LEFT JOIN dbo.[CAT DE ESTRUCTURA NIVEL 5] n5 ON mc.ID_N5 = n5.ID_N5",
       "LEFT JOIN dbo.[CAT DE ESTRUCTURA NIVEL 6] n6 ON mc.ID_N6 = n6.ID_N6",
       "LEFT JOIN dbo.[CAT DE ESTRUCTURA NIVEL 7] n7 ON mc.ID_N7 = n7.ID_N7",
       "LEFT JOIN dbo.[CAT DE ESTRUCTURA NIVEL 8] n8 ON mc.ID_N8 = n8.ID_N8"
    }

        ' Agregar condiciones fijas
        whereCondition &= " AND mp.CODIGO <> '0000000000' AND mc.ID_SITUACION = 1"

        ' Ejecutar la consulta principal y obtener resultados en dt
        Dim dt As DataTable = ConsultarDatos("dbo.[MAESTRO DE CARGOS] AS mc", joins, "mp.CODIGO", whereCondition, parametros, "", "mp.CODIGO ASC")

        ' Obtener los valores de CODIGO y guardarlos en una lista
        Dim listaCodigos As New List(Of String)
        For Each row As DataRow In dt.Rows
            If Not IsDBNull(row("CODIGO")) Then
                listaCodigos.Add("'" & row("CODIGO").ToString() & "'") ' Agregar comillas simples para la consulta SQL
            End If
        Next

        ' Si no hay códigos, evitar consulta con IN vacío
        If listaCodigos.Count = 0 Then
            MessageBox.Show("No se encontraron códigos de empleados.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return Nothing
        End If

        ' Obtener el último valor válido del ComboBox seleccionado
        Dim displayMemberSeleccionado As String = ObtenerUltimoDisplayMember()

        ' Construir la cláusula IN con los códigos obtenidos
        Dim codigosIn As String = String.Join(", ", listaCodigos)

        ' Obtener las fechas seleccionadas en los DateTimePicker
        Dim fechaInicio As String = DateTimePicker1.Value.ToString("yyyy-MM-dd")
        Dim fechaFin As String = DateTimePicker2.Value.ToString("yyyy-MM-dd")

        ' Definir la consulta secundaria con los códigos extraídos y las fechas dinámicas
        Dim whereCodigos As String = "mdm.FECHA_MARCA BETWEEN CAST('" & fechaInicio & "' AS DATE) AND CAST('" & fechaFin & "' AS DATE) " &
                             "AND mp.CODIGO IN (" & codigosIn & ")"


        ' Ejecutar la segunda consulta con los códigos extraídos
        Dim dtMarcas As DataTable = ConsultarDatos("dbo.[MAESTRO DE MARCAS] AS mdm",
        New List(Of String) From {
            "INNER JOIN dbo.[MAESTRO DE PERSONAS] mp ON mp.CODIGO = mdm.ID_EMPLEADO",
            "INNER JOIN dbo.[CAT DE RELOJ] CT ON CT.IDRELOJ = mdm.IDRELOJ"
        },
         "'" + displayMemberSeleccionado + "' AS DEPARTAMENTO, " &
        "mp.NOMBRES + ' ' + mp.APELLIDOS as NombreCompleto,
         mdm.ID_EMPLEADO,
         CT.DESCRIPCION,
         mdm.FECHA_MARCA,
        COALESCE(CONVERT(VARCHAR, MIN(CASE
                                         WHEN mdm.TIPO_MARCA = 'ENTRADA' THEN mdm.HORA_MARCA
                                     END), 120), 'USUARIO NO MARCO') AS ENTRADA,
       COALESCE(CONVERT(VARCHAR, MAX(CASE
                                         WHEN mdm.TIPO_MARCA = 'SALIDA' THEN mdm.HORA_MARCA
                                     END), 120), 'USUARIO NO MARCO') AS SALIDA",
        whereCodigos,
        New Dictionary(Of String, Object),
        "mdm.ID_EMPLEADO, CT.DESCRIPCION, mp.NOMBRES, mp.APELLIDOS, mdm.FECHA_MARCA",
        "mdm.FECHA_MARCA ASC, CT.DESCRIPCION, mdm.ID_EMPLEADO ASC"
    )

        ' Asignar la DataTable al DataGridView
        DataGridView1.DataSource = dtMarcas

        ' Cambiar los nombres de las columnas para la visualización
        With DataGridView1
            .Columns("DEPARTAMENTO").HeaderText = "DEPARTAMENTO"
            .Columns("NombreCompleto").HeaderText = "NOMBRE COMPLETO"
            .Columns("ID_EMPLEADO").HeaderText = "CODIGO EMPLEADO"
            .Columns("DESCRIPCION").HeaderText = "RELOJ"
            .Columns("FECHA_MARCA").HeaderText = "FECHA MARCA"
            .Columns("ENTRADA").HeaderText = "ENTRADA"
            .Columns("SALIDA").HeaderText = "SALIDA"
        End With
        Return dtMarcas
    End Function



    ''' <summary>
    ''' Método genérico para ejecutar consultas SQL con soporte para INNER JOIN, LEFT JOIN, alias, filtros dinámicos y ORDER BY.
    ''' </summary>
    ''' <param name="tabla">Tabla principal con alias opcional (ej: "Clientes c")</param>
    ''' <param name="joins">Lista de JOINs con alias y condiciones (ej: "INNER JOIN Direcciones d ON c.id = d.id_cliente")</param>
    ''' <param name="campos">Lista de campos a seleccionar (ej: "c.nombre, d.direccion")</param>
    ''' <param name="condiciones">Condiciones WHERE (ej: "c.edad > @edad AND d.ciudad = @ciudad")</param>
    ''' <param name="parametros">Diccionario con parámetros (ej: {"@edad", 30}, {"@ciudad", "Madrid"})</param>
    ''' <param name="orden">Campo(s) para ordenar (ej: "c.nombre ASC, d.direccion DESC")</param>
    ''' <returns>DataTable con los resultados</returns>
    Public Function ConsultarDatos(tabla As String, joins As List(Of String), campos As String, condiciones As String, parametros As Dictionary(Of String, Object), Optional groupby As String = "", Optional orden As String = "") As DataTable
        Dim dt As New DataTable()

        ' Validar que los parámetros obligatorios no estén vacíos
        If String.IsNullOrEmpty(tabla) OrElse String.IsNullOrEmpty(campos) Then
            MessageBox.Show("El nombre de la tabla y los campos son obligatorios.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return dt
        End If

        ' Construir la consulta dinámica
        Dim query As New Text.StringBuilder()
        query.Append($"SELECT {campos} FROM {tabla}")

        ' Agregar INNER JOIN / LEFT JOIN si existen
        If joins IsNot Nothing AndAlso joins.Count > 0 Then
            For Each joinTable In joins
                query.Append($" {joinTable}")
            Next
        End If

        ' Agregar condiciones WHERE
        If Not String.IsNullOrEmpty(condiciones) Then
            query.Append($" WHERE {condiciones}")
        End If

        ' Agregar ordenamiento ORDER BY
        If Not String.IsNullOrEmpty(groupby) Then
            query.Append($" GROUP BY {groupby}")
        End If

        ' Agregar ordenamiento ORDER BY
        If Not String.IsNullOrEmpty(orden) Then
            query.Append($" ORDER BY {orden}")
        End If

        Try
            ' Validar que la conexión está abierta
            If Conexion.CxRRHH.State = ConnectionState.Closed Then
                Conexion.CBD_RRHH()
            End If

            Using cmd As New SqlCommand(query.ToString(), Conexion.CxRRHH)
                ' Agregar parámetros a la consulta
                If parametros IsNot Nothing Then
                    For Each param In parametros
                        cmd.Parameters.AddWithValue(param.Key, param.Value)
                    Next
                End If

                ' Ejecutar la consulta
                Dim adapter As New SqlDataAdapter(cmd)
                adapter.Fill(dt)
            End Using
        Catch ex As Exception
            MessageBox.Show("Error en la consulta: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return dt
    End Function

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Call VERIFICAR_ACCESOS("376") : If HAY_ACCESO = False Then : Exit Sub : End If

        Dim displayMemberSeleccionado As String = ObtenerUltimoDisplayMember()
        MessageBox.Show("Último DisplayMember seleccionado: " & displayMemberSeleccionado, "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)

        ' Crear un cuadro de diálogo para seleccionar la ubicación de guardado
        Dim saveFileDialog As New SaveFileDialog()

        ' Configurar opciones del diálogo
        saveFileDialog.Filter = "Archivos de Excel (*.xlsx)|*.xlsx" ' Solo permite archivos .xlsx
        saveFileDialog.Title = "Guardar archivo de Excel"
        saveFileDialog.FileName = "DatosExportadosMarcadas.xlsx" ' Nombre predeterminado

        ' Mostrar el cuadro de diálogo y verificar si el usuario seleccionó una ubicación
        If saveFileDialog.ShowDialog() = DialogResult.OK Then
            Dim rutaArchivo As String = saveFileDialog.FileName

            ' Llamar a la función de la clase ExportadorExcelInterop
            ExportadorExcelInterop.ExportarAExcelOptimizado(DataGridView1, rutaArchivo)
        Else
            MessageBox.Show("Exportación cancelada.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Function ObtenerUltimoDisplayMember() As String
        ' Verificar si hay un último ComboBox seleccionado
        If ultimoComboBoxSeleccionado Is Nothing OrElse ultimoComboBoxSeleccionado.SelectedItem Is Nothing Then
            Return "Ninguno"
        End If

        ' Obtener la lista de ComboBox en orden
        Dim comboBoxes As ComboBox() = {ComboBox1, ComboBox2, ComboBox3, ComboBox4, ComboBox5, ComboBox6, ComboBox7, ComboBox8}

        ' Obtener el índice del último ComboBox seleccionado
        Dim indiceActual As Integer = Array.IndexOf(comboBoxes, ultimoComboBoxSeleccionado)

        ' Buscar el DisplayMember del último seleccionado
        Dim row As DataRowView = TryCast(ultimoComboBoxSeleccionado.SelectedItem, DataRowView)
        If row IsNot Nothing Then
            Dim displayValue As String = row(ultimoComboBoxSeleccionado.DisplayMember).ToString()

            ' Si el valor es "SELECCIONE", buscar el ComboBox anterior con una selección válida
            If displayValue = "SELECCIONE" Then
                For i As Integer = indiceActual - 1 To 0 Step -1
                    If comboBoxes(i).SelectedIndex > 0 Then ' Validar que no sea "SELECCIONE"
                        Dim previousRow As DataRowView = TryCast(comboBoxes(i).SelectedItem, DataRowView)
                        If previousRow IsNot Nothing Then
                            Return previousRow(comboBoxes(i).DisplayMember).ToString()
                        End If
                    End If
                Next
            Else
                Return displayValue ' Si el valor no es "SELECCIONE", devolverlo
            End If
        End If

        Return "Ninguno"
    End Function

End Class