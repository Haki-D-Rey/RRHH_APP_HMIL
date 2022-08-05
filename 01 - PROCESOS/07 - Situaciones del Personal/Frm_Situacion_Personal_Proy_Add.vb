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
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM  [dbo].[CAT DE SITUACION DE PERSONAL] WHERE ([ID_T_SIT_P] = 2) AND ([ID_SIT_P] = 12 OR [ID_SIT_P] = 17 OR [ID_SIT_P] = 25 OR [ID_SIT_P] = 22 OR [ID_SIT_P] = 19 OR [ID_SIT_P] = 5 OR [ID_SIT_P] = 7 OR [ID_SIT_P] = 16)  ORDER BY DESCRIPCION", Conexion.CxRRHH)
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
        xFECHA = Mid(Me.DateTimePicker1.Value, 1, 10)
        If id_EMPLEADO = 0 Then
            MsgBox("No se encuentra el codigo del Empleado", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        If Me.ComboBox1.Text = "POST TURNO 24" Then
            FE = Mid(Me.DateTimePicker1.Value, 1, 10)
            Call BUSCAR_SI_EXISTE_DIA_EN_SITUACION_DE_PERSONAL()
            If EXISTE_DIA_EN_SP = True Then
                Call EDITAR_REGISTRO_SITUACION_SELECCIONADA()
            Else
                Call AGREGAR_REGISTRO_DIA_SITUACION_SELECCIONADA()
            End If
            DIA_LIBRE_1 = CDate(xFECHA).AddDays(1)
            FE = DIA_LIBRE_1
            Call BUSCAR_SI_EXISTE_DIA_EN_SITUACION_DE_PERSONAL()
            If EXISTE_DIA_EN_SP = True Then
                Call EDITAR_REGISTRO_DIA_LIBRE()
            Else
                Call AGREGAR_REGISTRO_DIA_LIBRE()
            End If
            DIA_LIBRE_2 = CDate(xFECHA).AddDays(2)
            FE = DIA_LIBRE_2
            Call BUSCAR_SI_EXISTE_DIA_EN_SITUACION_DE_PERSONAL()
            If EXISTE_DIA_EN_SP = True Then
                Call EDITAR_REGISTRO_DIA_LIBRE()
            Else
                Call AGREGAR_REGISTRO_DIA_LIBRE()
            End If
        End If
        '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        If Me.ComboBox1.Text = "POST TURNO 12" Then
            FE = Mid(Me.DateTimePicker1.Value, 1, 10)
            Call BUSCAR_SI_EXISTE_DIA_EN_SITUACION_DE_PERSONAL()
            If EXISTE_DIA_EN_SP = True Then
                Call EDITAR_REGISTRO_SITUACION_SELECCIONADA()
            Else
                Call AGREGAR_REGISTRO_DIA_SITUACION_SELECCIONADA()
            End If
            DIA_LIBRE_1 = CDate(xFECHA).AddDays(1)
            FE = DIA_LIBRE_1
            Call BUSCAR_SI_EXISTE_DIA_EN_SITUACION_DE_PERSONAL()
            If EXISTE_DIA_EN_SP = True Then
                Call EDITAR_REGISTRO_DIA_LIBRE()
            Else
                Call AGREGAR_REGISTRO_DIA_LIBRE()
            End If
        End If
        '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        If Me.ComboBox1.Text = "PAGO DE TURNO" Then
            FE = Mid(Me.DateTimePicker1.Value, 1, 10)
            Call BUSCAR_SI_EXISTE_DIA_EN_SITUACION_DE_PERSONAL()
            If EXISTE_DIA_EN_SP = True Then
                Call EDITAR_REGISTRO_SITUACION_SELECCIONADA()
            Else
                Call AGREGAR_REGISTRO_DIA_SITUACION_SELECCIONADA()
            End If
        End If
        '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        If Me.ComboBox1.Text = "VACACIONES" Then
            Call RECORRER_DIAS()
        End If

        CERRAR = False
        Me.Close()
    End Sub
    Private Sub RECORRER_DIAS()
        Me.DateTimePicker10.Value = Mid(Me.DateTimePicker1.Value, 1, 10)
        Me.DateTimePicker20.Value = Mid(Me.DateTimePicker2.Value, 1, 10)
        Me.DateTimePicker30.Value = Mid(Me.DateTimePicker1.Value, 1, 10)
        Me.Cursor = Cursors.WaitCursor
        For I2 As Integer = 1 To DateDiff(DateInterval.Day, Me.DateTimePicker10.Value, Me.DateTimePicker20.Value.AddDays(1))
            Application.DoEvents()
            Me.DateTimePicker30.Refresh()
            FE = Mid(Me.DateTimePicker30.Value, 1, 10)
            Call BUSCAR_SI_EXISTE_DIA_EN_SITUACION_DE_PERSONAL()
            If EXISTE_DIA_EN_SP = True Then
                Call EDITAR_REGISTRO_SITUACION_SELECCIONADA()
            Else
                Call AGREGAR_REGISTRO_DIA_SITUACION_SELECCIONADA()
            End If
            Me.DateTimePicker30.Value = Me.DateTimePicker30.Value.AddDays(1)
            Me.DateTimePicker30.Refresh()
        Next
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub EDITAR_REGISTRO_SITUACION_SELECCIONADA()
        Call Clases.BUSCAR_FECHA_Y_HORA()
        Dim F2 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(FECHA_SERVIDOR, 1, 10) + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE PROYECCION] SET ID_M_P = " & id_EMPLEADO & ", ID_SIT_P = " & Me.ComboBox1.SelectedValue & ", FECHA_PROY = " & F2 & ", OBSERVACIONES = '" & Trim(Me.TextBox3.Text) & "'  WHERE ID_PROY = " & ID_PARA_EDITAR & ""
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
        Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + FE + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE PROYECCION] (ID_M_P, ID_SIT_P, FECHA_PROY, OBSERVACIONES)" &
                              "values (" & id_EMPLEADO & ", " & Me.ComboBox1.SelectedValue & ", " & F1 & ", '" & Me.TextBox3.Text & "')"
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
        Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + FE + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE PROYECCION] (ID_M_P, ID_SIT_P, FECHA_PROY, OBSERVACIONES)" &
                              "values (" & id_EMPLEADO & ", 29, " & F1 & ", '" & Me.TextBox3.Text & "')"
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
        Dim F2 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(FECHA_SERVIDOR, 1, 10) + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE PROYECCION] SET ID_M_P = " & id_EMPLEADO & ", ID_SIT_P = 29, FECHA_PROY = " & F2 & ", OBSERVACIONES = '" & Trim(Me.TextBox3.Text) & "'  WHERE ID_PROY = " & ID_PARA_EDITAR & ""
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
        If Me.ComboBox1.Text = "VACACIONES" Then
            Me.DateTimePicker2.Enabled = True
        Else
            Me.DateTimePicker2.Enabled = False
        End If
    End Sub
    Private Sub BUSCAR_SI_EXISTE_DIA_EN_SITUACION_DE_PERSONAL()
        cEMPLEADO = Format(CInt(cEMPLEADO), "0000000000")
        Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + FE + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable3 As New DataTable
        Dim MiDataAdapter3 As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter3 = New SqlDataAdapter("SELECT * FROM [VISTA MAESTRO DE PROYECCION] WHERE (ID_M_P = " & id_EMPLEADO & ") AND (FECHA_PROY = " & F1 & ") ORDER BY FECHA_PROY", Conexion.CxRRHH)
            MiDataTable3.Clear()
            MiDataAdapter3.Fill(MiDataTable3)
            If MiDataTable3.Rows.Count <> 0 Then
                EXISTE_DIA_EN_SP = True
                ID_PARA_EDITAR = MiDataTable3.Rows(0).Item(0).ToString
            Else
                EXISTE_DIA_EN_SP = False
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
End Class