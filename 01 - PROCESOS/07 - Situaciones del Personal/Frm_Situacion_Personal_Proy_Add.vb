Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Situacion_Personal_Proy_Add
    Private Sub Frm_Situacion_Personal_Proy_Add_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox3.Text = ""
            Me.ComboBox1.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Situacion_Personal_Proy_Add_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call Clases.BUSCAR_FECHA_Y_HORA()
        Me.DateTimePicker1.Value = FECHA_SERVIDOR
        Call CARGAR_COMBO_CAT_DE_SITUACION_DE_PERSONAL()
        Me.ComboBox1.Text = "POST TURNO 24"
        Me.TextBox3.Text = ""
        Me.ComboBox1.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.TextBox3.Text = ""
        Me.ComboBox1.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_SITUACION_DE_PERSONAL()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM  [dbo].[CAT DE SITUACION DE PERSONAL] WHERE ([ID_T_SIT_P] = 2) AND [ID_SIT_P] IN(12, 17, 25, 22, 19 ,5 ,7 ,16, 8)  ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_SIT_P"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim DATO_1 As Boolean
    Private Sub BUSCAR_DATO_1()
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
                id_EMPLEADO = MiDataTable.Rows(0).Item(0).ToString
                DATO_1 = True
            Else
                id_EMPLEADO = 0
                DATO_1 = False
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub DateTimePicker1_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub DateTimePicker2_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Dim DIA_LIBRE_1, DIA_LIBRE_2 As String
    Dim EXISTE_DIA_EN_SP As Boolean
    Dim FE As String
    Dim ID_PARA_EDITAR As Integer
    Dim xFECHA As String

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim fechaDesde As Date = Me.DateTimePicker1.Value
        Dim fechaHasta As Date = Me.DateTimePicker2.Value

        If id_EMPLEADO = 0 Then
            MsgBox("No se encuentra el código del Empleado", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If

        ' Verifica el turno seleccionado y asigna los días libres correspondientes
        Select Case Me.ComboBox1.Text
            Case "POST TURNO 24"
                ProcesarTurno(fechaDesde, fechaHasta, diasLibresPorTurno:=2)

            Case "POST TURNO 12"
                ProcesarTurno(fechaDesde, fechaHasta, diasLibresPorTurno:=1)

            Case "PAGO DE TURNO"
                ProcesarPagoDeTurno(fechaDesde, fechaHasta)

            Case Else
                Call RECORRER_DIAS()
        End Select

        CERRAR = False
        Me.Close()
    End Sub

    ' Procesa el turno para POST TURNO 24 y POST TURNO 12, ajustando los días libres según el día de inicio
    Private Sub ProcesarTurno(fechaDesde As Date, fechaHasta As Date, diasLibresPorTurno As Integer)
        Dim diasLibres As Integer

        ' Ajusta los días libres según el día de inicio
        Select Case fechaDesde.DayOfWeek
            Case DayOfWeek.Friday
                diasLibres = diasLibresPorTurno
            Case DayOfWeek.Saturday
                diasLibres = Math.Max(diasLibresPorTurno - 1, 0)
            Case DayOfWeek.Sunday
                diasLibres = 0
        End Select

        ' Procesa el primer día del turno
        FE = fechaDesde.ToString("yyyy-MM-dd")
        Call BUSCAR_SI_EXISTE_DIA_EN_SITUACION_DE_PERSONAL()
        If EXISTE_DIA_EN_SP Then
            Call EDITAR_REGISTRO_SITUACION_SELECCIONADA()
        Else
            Call AGREGAR_REGISTRO_DIA_SITUACION_SELECCIONADA()
        End If

        ' Procesa los días libres, si aplica
        For i As Integer = 1 To diasLibres
            FE = fechaDesde.AddDays(i).ToString("yyyy-MM-dd")
            Call BUSCAR_SI_EXISTE_DIA_EN_SITUACION_DE_PERSONAL()
            If EXISTE_DIA_EN_SP Then
                Call EDITAR_REGISTRO_DIA_LIBRE()
            Else
                Call AGREGAR_REGISTRO_DIA_LIBRE()
            End If
        Next
    End Sub

    ' Procesa el turno de pago
    Private Sub ProcesarPagoDeTurno(fechaDesde As Date, fechaHasta As Date)
        FE = fechaDesde.ToString("yyyy-MM-dd")
        Call BUSCAR_SI_EXISTE_DIA_EN_SITUACION_DE_PERSONAL()
        If EXISTE_DIA_EN_SP Then
            Call EDITAR_REGISTRO_SITUACION_SELECCIONADA()
        Else
            Call AGREGAR_REGISTRO_DIA_SITUACION_SELECCIONADA()
        End If
    End Sub

    Private Sub RECORRER_DIAS()
        ' Establecer las fechas desde y hasta.
        Dim fechaDesde As Date = Me.DateTimePicker1.Value.Date
        Dim fechaHasta As Date = Me.DateTimePicker2.Value.Date

        Me.Cursor = Cursors.WaitCursor

        Dim currentDate As Date = fechaDesde

        While currentDate <= fechaHasta

            FE = currentDate.ToString("yyyy-MM-dd")

            ' Verificar si el día ya existe en la situación de personal.
            Call BUSCAR_SI_EXISTE_DIA_EN_SITUACION_DE_PERSONAL()

            ' Si existe, editar el registro, de lo contrario, agregar un nuevo registro.
            If EXISTE_DIA_EN_SP Then
                Call EDITAR_REGISTRO_SITUACION_SELECCIONADA()
            Else
                Call AGREGAR_REGISTRO_DIA_SITUACION_SELECCIONADA()
            End If

            ' Refrescar el control DateTimePicker30, si es necesario (opcional si no se usa).
            Me.DateTimePicker30.Value = currentDate
            Me.DateTimePicker30.Refresh()

            ' Asegurar que se procesen los eventos de la interfaz de usuario.
            Application.DoEvents()

            ' Incrementar currentDate por un día.
            currentDate = currentDate.AddDays(1)
        End While

        ' Restaurar el cursor a su estado predeterminado.
        Me.Cursor = Cursors.Default
    End Sub

    'Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
    '    xFECHA = Mid(Me.DateTimePicker1.Value, 1, 10)
    '    If id_EMPLEADO = 0 Then
    '        MsgBox("No se encuentra el codigo del Empleado", vbInformation, "Mensaje del Sistema")
    '        Exit Sub
    '    End If
    '    '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '    If Me.ComboBox1.Text = "POST TURNO 24" Then
    '        FE = Mid(Me.DateTimePicker1.Value, 1, 10)
    '        Call BUSCAR_SI_EXISTE_DIA_EN_SITUACION_DE_PERSONAL()
    '        If EXISTE_DIA_EN_SP = True Then
    '            Call EDITAR_REGISTRO_SITUACION_SELECCIONADA()
    '        Else
    '            Call AGREGAR_REGISTRO_DIA_SITUACION_SELECCIONADA()
    '        End If
    '        DIA_LIBRE_1 = CDate(xFECHA).AddDays(1)
    '        FE = DIA_LIBRE_1
    '        Call BUSCAR_SI_EXISTE_DIA_EN_SITUACION_DE_PERSONAL()
    '        If EXISTE_DIA_EN_SP = True Then
    '            Call EDITAR_REGISTRO_DIA_LIBRE()
    '        Else
    '            Call AGREGAR_REGISTRO_DIA_LIBRE()
    '        End If
    '        DIA_LIBRE_2 = CDate(xFECHA).AddDays(2)
    '        FE = DIA_LIBRE_2
    '        Call BUSCAR_SI_EXISTE_DIA_EN_SITUACION_DE_PERSONAL()
    '        If EXISTE_DIA_EN_SP = True Then
    '            Call EDITAR_REGISTRO_DIA_LIBRE()
    '        Else
    '            Call AGREGAR_REGISTRO_DIA_LIBRE()
    '        End If
    '    End If
    '    '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '    If Me.ComboBox1.Text = "POST TURNO 12" Then
    '        FE = Mid(Me.DateTimePicker1.Value, 1, 10)
    '        Call BUSCAR_SI_EXISTE_DIA_EN_SITUACION_DE_PERSONAL()
    '        If EXISTE_DIA_EN_SP = True Then
    '            Call EDITAR_REGISTRO_SITUACION_SELECCIONADA()
    '        Else
    '            Call AGREGAR_REGISTRO_DIA_SITUACION_SELECCIONADA()
    '        End If
    '        DIA_LIBRE_1 = CDate(xFECHA).AddDays(1)
    '        FE = DIA_LIBRE_1
    '        Call BUSCAR_SI_EXISTE_DIA_EN_SITUACION_DE_PERSONAL()
    '        If EXISTE_DIA_EN_SP = True Then
    '            Call EDITAR_REGISTRO_DIA_LIBRE()
    '        Else
    '            Call AGREGAR_REGISTRO_DIA_LIBRE()
    '        End If
    '    End If
    '    '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '    If Me.ComboBox1.Text = "PAGO DE TURNO" Then
    '        FE = Mid(Me.DateTimePicker1.Value, 1, 10)
    '        Call BUSCAR_SI_EXISTE_DIA_EN_SITUACION_DE_PERSONAL()
    '        If EXISTE_DIA_EN_SP = True Then
    '            Call EDITAR_REGISTRO_SITUACION_SELECCIONADA()
    '        Else
    '            Call AGREGAR_REGISTRO_DIA_SITUACION_SELECCIONADA()
    '        End If
    '    End If
    '    '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '    If Me.ComboBox1.Text = "VACACIONES" Then
    '        Call RECORRER_DIAS()
    '    End If

    '    CERRAR = False
    '    Me.Close()
    'End Sub
    'Private Sub RECORRER_DIAS()
    '    Me.DateTimePicker10.Value = Mid(Me.DateTimePicker1.Value, 1, 10)
    '    Me.DateTimePicker20.Value = Mid(Me.DateTimePicker2.Value, 1, 10)
    '    Me.DateTimePicker30.Value = Mid(Me.DateTimePicker1.Value, 1, 10)
    '    Me.Cursor = Cursors.WaitCursor
    '    For I2 As Integer = 1 To DateDiff(DateInterval.Day, Me.DateTimePicker10.Value, Me.DateTimePicker20.Value.AddDays(1))
    '        Application.DoEvents()
    '        Me.DateTimePicker30.Refresh()
    '        FE = Mid(Me.DateTimePicker30.Value, 1, 10)
    '        Call BUSCAR_SI_EXISTE_DIA_EN_SITUACION_DE_PERSONAL()
    '        If EXISTE_DIA_EN_SP = True Then
    '            Call EDITAR_REGISTRO_SITUACION_SELECCIONADA()
    '        Else
    '            Call AGREGAR_REGISTRO_DIA_SITUACION_SELECCIONADA()
    '        End If
    '        Me.DateTimePicker30.Value = Me.DateTimePicker30.Value.AddDays(1)
    '        Me.DateTimePicker30.Refresh()
    '    Next
    '    Me.Cursor = Cursors.Default
    'End Sub
    Private Sub EDITAR_REGISTRO_SITUACION_SELECCIONADA()
        ' Cambiar la fecha a DateTime, obteniendo la fecha actual del servidor (o la fecha que necesites)
        Dim F2 As DateTime = DateTime.Parse(FE) ' Suponiendo que FECHA_SERVIDOR es una cadena de fecha
        Dim Valor As String = Me.TextBox1.Text

        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            ' Aquí usamos F2 como DateTime y lo pasamos al SQL sin formatear a cadena
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE PROYECCION] SET ID_M_P = " & id_EMPLEADO & ", ID_SIT_P = " & Me.ComboBox1.SelectedValue & ", VALOR_SIT = @Valor , FECHA_PROY = @FechaProy, OBSERVACIONES = '" & Trim(Me.TextBox3.Text) & "'  WHERE ID_PROY = " & ID_PARA_EDITAR & ""
            MiSqlCommand.Parameters.AddWithValue("@FechaProy", F2) ' Usamos el parámetro para evitar problemas de formato
            MiSqlCommand.Parameters.AddWithValue("@Valor", Valor)
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub

    Private Sub AGREGAR_REGISTRO_DIA_SITUACION_SELECCIONADA()
        ' Asegurarse de que FE esté en formato DateTime
        Dim F1 As DateTime = DateTime.Parse(FE) ' Suponiendo que FE es una cadena de fecha
        Dim Valor As String = Me.TextBox1.Text
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            ' Usamos F1 como DateTime y pasamos al SQL mediante parámetros
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE PROYECCION] (ID_M_P, ID_SIT_P, FECHA_PROY, OBSERVACIONES, VALOR_SIT) " &
                                  "values (@IdEmpleado, @IdSitP, @FechaProy, @Observaciones, @Valor)"
            MiSqlCommand.Parameters.AddWithValue("@IdEmpleado", id_EMPLEADO)
            MiSqlCommand.Parameters.AddWithValue("@IdSitP", Me.ComboBox1.SelectedValue)
            MiSqlCommand.Parameters.AddWithValue("@FechaProy", F1)  ' Usamos el parámetro para la fecha
            MiSqlCommand.Parameters.AddWithValue("@Observaciones", Me.TextBox3.Text)
            MiSqlCommand.Parameters.AddWithValue("@Valor", Valor)
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub

    Private Sub AGREGAR_REGISTRO_DIA_LIBRE()
        ' Asegúrate de que FE esté en formato DateTime
        Dim F1 As DateTime = DateTime.Parse(FE)  ' Suponiendo que FE es una cadena de fecha
        Dim Valor As String = Me.TextBox1.Text
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            ' Usamos F1 como DateTime y lo pasamos al SQL mediante parámetros
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE PROYECCION] (ID_M_P, ID_SIT_P, FECHA_PROY, OBSERVACIONES, VALOR_SIT) " &
                                  "values (@IdEmpleado, 29, @FechaProy, @Observaciones, @Valor)"
            MiSqlCommand.Parameters.AddWithValue("@IdEmpleado", id_EMPLEADO)
            MiSqlCommand.Parameters.AddWithValue("@FechaProy", F1)  ' Usamos el parámetro para la fecha
            MiSqlCommand.Parameters.AddWithValue("@Observaciones", Me.TextBox3.Text)
            MiSqlCommand.Parameters.AddWithValue("@Valor", Valor)
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub

    Private Sub EDITAR_REGISTRO_DIA_LIBRE()
        Call Clases.BUSCAR_FECHA_Y_HORA()
        ' Asegurarse de que FECHA_SERVIDOR esté en formato DateTime
        Dim F2 As DateTime = DateTime.Parse(FE)
        Dim Valor As String = Me.TextBox1.Text
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            ' Usamos F2 como DateTime y lo pasamos al SQL mediante parámetros
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE PROYECCION] SET ID_M_P = @IdEmpleado, ID_SIT_P = 29, FECHA_PROY = @FechaProy, OBSERVACIONES = @Observaciones, VALOR_SIT =  WHERE ID_PROY = @IdProy"
            MiSqlCommand.Parameters.AddWithValue("@IdEmpleado", id_EMPLEADO)
            MiSqlCommand.Parameters.AddWithValue("@FechaProy", F2)  ' Usamos el parámetro para la fecha
            MiSqlCommand.Parameters.AddWithValue("@Observaciones", Trim(Me.TextBox3.Text))
            MiSqlCommand.Parameters.AddWithValue("@Valor", Trim(Me.TextBox1.Text))
            MiSqlCommand.Parameters.AddWithValue("@IdProy", ID_PARA_EDITAR)
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim selectedRow As DataRowView = DirectCast(ComboBox1.SelectedItem, DataRowView)
        Dim selectedValue As Integer = Convert.ToInt32(selectedRow("ID_SIT_P"))
        If {5, 7, 19, 22, 25, 12, 16}.Contains(selectedValue) Then
            Me.DateTimePicker2.Enabled = False
        Else
            Me.DateTimePicker2.Enabled = True
        End If

        'Dim selectedRow As DataRowView = DirectCast(ComboBox1.SelectedItem, DataRowView)

        'Dim selectedValue As Integer = Convert.ToInt32(selectedRow("ID_SIT_P"))
        'If {5, 7, 19, 22, 25, 12, 16}.Contains(selectedValue) Then
        '    MonthCalendar1.Visible = True
        '    MonthCalendar1.Size = New Size(305, 200) ' Tamaño cuando el calendario está visible
        '    Me.Size = New Size(630, 437) ' Tamaño del formulario cuando el calendario está visible

        '    ' Ocultar los otros controles
        '    Label7.Visible = False
        '    DateTimePicker1.Visible = False
        '    Label2.Visible = False
        '    DateTimePicker2.Visible = False
        'Else
        '    ' Ocultar el MonthCalendar y ajustar el tamaño
        '    MonthCalendar1.Visible = False
        '    MonthCalendar1.Size = New Size(337, 36) ' Tamaño cuando el calendario está oculto
        '    Me.Size = New Size(630, 200) ' Tamaño del formulario cuando el calendario está oculto

        '    ' Mostrar los otros controles
        '    Label7.Visible = True
        '    DateTimePicker1.Visible = True
        '    Label2.Visible = True
        '    DateTimePicker2.Visible = True
        'End If
    End Sub

    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub BUSCAR_SI_EXISTE_DIA_EN_SITUACION_DE_PERSONAL()
        cEMPLEADO = Format(CInt(cEMPLEADO), "0000000000")
        ' Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + FE + "', 105), 23)"
        Dim F1 As String = FE

        ' Check if connection is open
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If

        ' Prepare the SQL query with parameters to prevent SQL injection
        Dim query As String = "SELECT * FROM [VISTA MAESTRO DE PROYECCION] WHERE (ID_M_P = @id_EMPLEADO) AND (FECHA_PROY = @fecha_PROY) ORDER BY FECHA_PROY"

        ' Create the DataTable and DataAdapter
        Dim MiDataTable3 As New DataTable()
        Dim MiDataAdapter3 As SqlDataAdapter

        Try
            ' Open the connection and fill the DataTable
            Conexion.ABD_RRHH()
            MiDataAdapter3 = New SqlDataAdapter(query, Conexion.CxRRHH)

            ' Add parameters to the query to prevent SQL injection
            MiDataAdapter3.SelectCommand.Parameters.AddWithValue("@id_EMPLEADO", id_EMPLEADO)
            MiDataAdapter3.SelectCommand.Parameters.AddWithValue("@fecha_PROY", F1)

            MiDataTable3.Clear()
            MiDataAdapter3.Fill(MiDataTable3)

            ' Check if any rows were returned
            If MiDataTable3.Rows.Count > 0 Then
                EXISTE_DIA_EN_SP = True
                ID_PARA_EDITAR = CInt(MiDataTable3.Rows(0).Item(0).ToString())
            Else
                EXISTE_DIA_EN_SP = False
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            ' Ensure the connection is properly closed
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub


End Class