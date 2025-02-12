Imports System.Data.SqlClient
Imports System.Reflection

Public Class Frm_Marcas

    ' Diccionario global para almacenar los valores seleccionados por cada ComboBox
    Private seleccionDatos As New Dictionary(Of String, Dictionary(Of String, String))

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
        Dim dt As DataTable = ConsultarDatos("dbo.[CAT DE ESTRUCTURA NIVEL 1]", joins, "*", "", parametros, "ID_N1 ASC")

        ' Rellenar el ComboBox1 con los resultados
        ConfigurarComboBoxPorDefecto()
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

        ' Buscar el ComboBox que activó el evento
        For i As Integer = 0 To comboBoxes.Length - 1
            If sender Is comboBoxes(i) Then
                ' Habilitar el CheckBox correspondiente si la selección no es "SELECCIONE"
                checkBoxes(i).Enabled = (comboBoxes(i).SelectedIndex > 0)
                checkBoxes(i).Checked = (comboBoxes(i).SelectedIndex > 0)
            End If
        Next
    End Sub

    Private Sub CheckBox_CheckedChanged(sender As Object, e As EventArgs)
        Dim comboBoxes As ComboBox() = {ComboBox1, ComboBox2, ComboBox3, ComboBox4, ComboBox5, ComboBox6, ComboBox7, ComboBox8}
        Dim checkBoxes As CheckBox() = {CheckBox1, CheckBox2, CheckBox3, CheckBox4, CheckBox5, CheckBox6, CheckBox7, CheckBox8}
        Dim keyMapping As String() = {"n1.SIGLAS", "n2.SIGLAS", "n3.SIGLAS", "n4.SIGLAS", "n5.SIGLAS", "n6.SIGLAS", "n7.SIGLAS", "n8.SIGLAS"} ' Llaves dinámicas

        ' Buscar el CheckBox que activó el evento
        For i As Integer = 0 To checkBoxes.Length - 1
            If sender Is checkBoxes(i) AndAlso checkBoxes(i).Checked Then
                ' Obtener valores del ComboBox asociado
                Dim valorTexto As String = ""
                Dim valorInterno As String = ""
                Dim valorIDN As String = keyMapping(i) ' ID_N asociado a este ComboBox

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

                ' Construir el WHERE y los parámetros después de actualizar la selección
                Dim whereCondition As String = ConstruirWhereCondition()
                Dim parametros As Dictionary(Of String, Object) = ConstruirParametros()


                ' Llamar a la consulta y mostrar los resultados en DataGridView1
                ActualizarDataGridView(whereCondition, parametros, valorInterno)
            End If
        Next
    End Sub



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

    Private Sub ActualizarDataGridView(whereCondition As String, parametros As Dictionary(Of String, Object), valor As String)
        ' Definir parámetros de la consulta
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

        Dim joins As New List(Of String) From {
           "INNER JOIN dbo.[CAT DE TIPO DE ESTRUCTURA] tes ON mc.ID_T_ES = tes.ID_T_ES",
           "INNER Join dbo.[CAT DE CARGOS DE ESTRUCTURA] ces ON mc.ID_CARGO_ES = ces.ID_CARGO_ES",
           "INNER Join dbo.[CAT DE GRADO PLANTILLA] gp ON mc.ID_GP = gp.ID_GP",
           "INNER Join dbo.[CAT DE GRADO REAL] gr ON mc.ID_GR = gr.ID_GR",
           "INNER Join dbo.[CAT DE CATEGORIA SALARIAL] cs ON mc.ID_CAT_SALARIAL = cs.ID_CAT_SALARIAL",
           "INNER Join dbo.[MAESTRO DE PERSONAS] mp ON mc.ID_M_P = mp.ID_M_P",
           "INNER Join dbo.[CAT DE ESTABLECIMIENTO] est ON mc.ID_ESTABLECIMIENTO = est.ID_ESTABLECIMIENTO",
           "INNER Join dbo.[CAT DE ADMINISTRADORES] adm ON mp.ID_ADMIN = adm.ID_ADMIN",
           "Left Join dbo.[CAT DE ESTRUCTURA NIVEL 1] n1 ON mc.ID_N1 = n1.ID_N1",
           "Left Join dbo.[CAT DE ESTRUCTURA NIVEL 2] n2 ON mc.ID_N2 = n2.ID_N2",
           "Left Join dbo.[CAT DE ESTRUCTURA NIVEL 3] n3 ON mc.ID_N3 = n3.ID_N3",
           "Left Join dbo.[CAT DE ESTRUCTURA NIVEL 4] n4 ON mc.ID_N4 = n4.ID_N4",
           "Left Join dbo.[CAT DE ESTRUCTURA NIVEL 5] n5 ON mc.ID_N5 = n5.ID_N5",
           "Left Join dbo.[CAT DE ESTRUCTURA NIVEL 6] n6 ON mc.ID_N6 = n6.ID_N6",
           "Left Join dbo.[CAT DE ESTRUCTURA NIVEL 7] n7 ON mc.ID_N7 = n7.ID_N7",
           "Left Join dbo.[CAT DE ESTRUCTURA NIVEL 8] n8 ON mc.ID_N8 = n8.ID_N8"
        }

        whereCondition = whereCondition & " AND mp.CODIGO <> '0000000000' AND mc.ID_SITUACION = 1"

        ' Ejecutar la consulta con filtro dinámico
        Dim dt As DataTable = ConsultarDatos("dbo.[MAESTRO DE CARGOS] AS mc", joins, "mc.ID_M_C, mc.ORDEN,
                                                n5.DESCRIPCION as AREA_DEPARTAMENTO,
                                                mc.ID_T_ES, tes.DESCRIPCION AS D_T_ES,
                                                mp.APELLIDOS + N' ' + mp.NOMBRES AS AN,
                                                mc.N_ORDEN_MIXTA, mc.N_ORDEN_MILITAR, mc.N_ORDEN_PAME, 
                                                mc.ID_CARGO_ES, ces.DESCRIPCION AS D_CARGO_ES,
                                                mc.ID_M_P, mp.CODIGO,
                                                mc.ID_SITUACION, mc.JEFE, mc.MIXTA, mc.MILITAR, mc.PAME, 
                                                mc.ID_ESTABLECIMIENTO, est.DESCRIPCION AS D_ESTABLECIMIENTO, 
                                                adm.DESCRIPCION AS D_ADMIN", whereCondition, parametros, "mp.CODIGO ASC")

        ' Mostrar los resultados en DataGridView1
        DataGridView1.DataSource = dt
    End Sub


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
    Public Function ConsultarDatos(tabla As String, joins As List(Of String), campos As String, condiciones As String, parametros As Dictionary(Of String, Object), Optional orden As String = "") As DataTable
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


End Class