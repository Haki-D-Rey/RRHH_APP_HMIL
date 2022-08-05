Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class xxx_Frm_Situacion_Personal_Proyectar
    Private Sub Frm_Situacion_Personal_Proyectar_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = ""
            Me.TextBox2.Text = ""
            Me.TextBox3.Text = ""
            Me.TextBox1.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Situacion_Personal_Proyectar_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox1.Text = Frm_Situacion_Personal.TextBox2.Text
        Call Clases.BUSCAR_FECHA_Y_HORA()
        Me.DateTimePicker1.Value = FECHA_SERVIDOR
        Call CARGAR_COMBO_CAT_DE_SITUACION_DE_PERSONAL()
        Me.ComboBox1.Text = "POST TURNO 24"
        Me.TextBox3.Text = ""
        Me.TextBox1.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox1.Focus()
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
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM  [dbo].[CAT DE SITUACION DE PERSONAL] WHERE ([ID_T_SIT_P] = 2) AND ([ID_SIT_P] = 12 OR [ID_SIT_P] = 17 OR [ID_SIT_P] = 25)  ORDER BY DESCRIPCION", Conexion.CxRRHH)
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
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            cEMPLEADO = CInt(Val(TextBox1.Text))
            TextBox1.Text = Format(CInt(cEMPLEADO), "0000000000")
            If Len(Me.TextBox1.Text) = 0 Or (Me.TextBox1.Text = "0000000000") Then
                MsgBox("El código que se digitó no es válido", vbInformation, "Mensaje del Sistema")
                Me.TextBox2.Focus()
                Exit Sub
            Else
                If isuID_TIPO_USUARIO = 5 Then
                    CADENAsql = "SELECT ID_M_P, CODIGO, AN FROM [dbo].[VISTA MAESTRO DE PERSONAS] WHERE CODIGO = '" & Me.TextBox1.Text & "' AND ID_ESTADO_P = 1 AND D_ADMIN = '" & isuUSUARIO & "'"
                Else
                    CADENAsql = "SELECT ID_M_P, CODIGO, AN FROM [dbo].[VISTA MAESTRO DE PERSONAS] WHERE CODIGO = '" & Me.TextBox1.Text & "' AND ID_ESTADO_P = 1"
                End If
                Call BUSCAR_DATO_1()
                If DATO_1 = True Then
                    Me.DateTimePicker1.Focus()
                Else
                    MsgBox("El código que se digitó no es válido o no se encuentra Asignado al usuario Actual", vbInformation, "Mensaje del Sistema")
                    Me.TextBox1.SelectAll()
                    Me.TextBox1.Focus()
                    Exit Sub
                End If
            End If
        End If
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
                Me.TextBox2.Text = MiDataTable.Rows(0).Item(2).ToString
                DATO_1 = True
            Else
                id_EMPLEADO = 0
                Me.TextBox2.Text = ""
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
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
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
            Me.TextBox1.SelectAll()
            Me.TextBox1.Focus()
            Exit Sub
        End If
        '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        If Me.ComboBox1.Text = "POST TURNO 24" Then
            'AGREGAR SITUACION SELECCIONADA
            FE = Mid(Me.DateTimePicker1.Value, 1, 10)
            Call BUSCAR_SI_EXISTE_DIA_EN_SITUACION_DE_PERSONAL()
            If EXISTE_DIA_EN_SP = True Then
                Call EDITAR_REGISTRO_SITUACION_SELECCIONADA()
            Else
                Call AGREGAR_REGISTRO_DIA_SITUACION_SELECCIONADA()
            End If
            'AGREGAR DIA LIBRE
            DIA_LIBRE_1 = CDate(xFECHA).AddDays(1)
            FE = DIA_LIBRE_1
            Call BUSCAR_SI_EXISTE_DIA_EN_SITUACION_DE_PERSONAL()
            If EXISTE_DIA_EN_SP = True Then
                Call EDITAR_REGISTRO_DIA_LIBRE()
            Else
                Call AGREGAR_REGISTRO_DIA_LIBRE()
            End If
            'AGREGAR DIA LIBRE
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
            'AGREGAR SITUACION SELECCIONADA
            FE = Mid(Me.DateTimePicker1.Value, 1, 10)
            Call BUSCAR_SI_EXISTE_DIA_EN_SITUACION_DE_PERSONAL()
            If EXISTE_DIA_EN_SP = True Then
                Call EDITAR_REGISTRO_SITUACION_SELECCIONADA()
            Else
                Call AGREGAR_REGISTRO_DIA_SITUACION_SELECCIONADA()
            End If
            'AGREGAR DIA LIBRE
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
        If Me.ComboBox1.Text = "VACACIONES" Then
            Call RECORRER_DIAS()
        End If
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
                Call EDITAR_REGISTRO_VACACIONES()
            Else
                Call AGREGAR_REGISTRO_VACACIONES()
            End If
            Me.DateTimePicker30.Value = Me.DateTimePicker30.Value.AddDays(1)
            Me.DateTimePicker30.Refresh()
        Next
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub EDITAR_REGISTRO_VACACIONES()
        Dim F2 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker30.Value, 1, 10) + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE SITUACION DE PERSONAL] SET ID_T_SIT_P = 2, ID_SIT_P = 17, VALOR_SIT = -1, OBSERVACIONES = '" & Trim(Me.TextBox3.Text) & "', DETALLE_1 = NULL , DETALLE_2= NULL, DETALLE_3 = NULL, USUARIO_ACT = '" & isuUSUARIO & "', FECHA_ACT = " & F2 & "  WHERE ID_M_A = " & ID_PARA_EDITAR & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub AGREGAR_REGISTRO_VACACIONES()
        Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker30.Value, 1, 10) + "', 105), 23)"
        Call Clases.BUSCAR_FECHA_Y_HORA()
        Dim F2 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(FECHA_SERVIDOR, 1, 10) + "', 105), 23)"
        ANTxDIA = 0.0833333333333333
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE SITUACION DE PERSONAL] (ID_M_P, DIA, ID_T_SIT_P, ID_SIT_P, VALOR_SIT, VALOR_DIA, OBSERVACIONES, DETALLE_1, DETALLE_2, DETALLE_3, USUARIO_ACT, FECHA_ACT)" &
                              "values (" & id_EMPLEADO & ", " & F1 & ", 2, 17, -1, '" & CDbl(ANTxDIA) & "','" & Trim(Me.TextBox3.Text) & "', NULL, NULL, NULL, '" & isuUSUARIO & "', " & F2 & ")"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
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
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE SITUACION DE PERSONAL] SET ID_T_SIT_P = 2, ID_SIT_P = " & Me.ComboBox1.SelectedValue & ", VALOR_SIT = 0, OBSERVACIONES = '" & Trim(Me.TextBox3.Text) & "', DETALLE_1 = NULL , DETALLE_2= NULL, DETALLE_3 = NULL, USUARIO_ACT = '" & isuUSUARIO & "', FECHA_ACT = " & F2 & "  WHERE ID_M_A = " & ID_PARA_EDITAR & ""
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
        Call Clases.BUSCAR_FECHA_Y_HORA()
        Dim F2 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(FECHA_SERVIDOR, 1, 10) + "', 105), 23)"
        ANTxDIA = 0.0833333333333333
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE SITUACION DE PERSONAL] (ID_M_P, DIA, ID_T_SIT_P, ID_SIT_P, VALOR_SIT, VALOR_DIA, OBSERVACIONES, DETALLE_1, DETALLE_2, DETALLE_3, USUARIO_ACT, FECHA_ACT)" &
                              "values (" & id_EMPLEADO & ", " & F1 & ", 2, " & Me.ComboBox1.SelectedValue & ", 0, '" & CDbl(ANTxDIA) & "','" & Trim(Me.TextBox3.Text) & "', NULL, NULL, NULL, '" & isuUSUARIO & "', " & F2 & ")"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim ANTxDIA As Double
    Private Sub AGREGAR_REGISTRO_DIA_LIBRE()
        Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + FE + "', 105), 23)"
        Call Clases.BUSCAR_FECHA_Y_HORA()
        Dim F2 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(FECHA_SERVIDOR, 1, 10) + "', 105), 23)"
        ANTxDIA = 0.0833333333333333
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE SITUACION DE PERSONAL] (ID_M_P, DIA, ID_T_SIT_P, ID_SIT_P, VALOR_SIT, VALOR_DIA, OBSERVACIONES, DETALLE_1, DETALLE_2, DETALLE_3, USUARIO_ACT, FECHA_ACT)" &
                              "values (" & id_EMPLEADO & ", " & F1 & ", 2, 29, 0, '" & CDbl(ANTxDIA) & "','" & Trim(Me.TextBox3.Text) & "', NULL, NULL, NULL, '" & isuUSUARIO & "', " & F2 & ")"
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
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE SITUACION DE PERSONAL] SET ID_T_SIT_P = 2, ID_SIT_P = 29, VALOR_SIT = 0, OBSERVACIONES = '" & Trim(Me.TextBox3.Text) & "', DETALLE_1 = NULL, DETALLE_2 = NULL, DETALLE_3 = NULL, USUARIO_ACT = '" & isuUSUARIO & "', FECHA_ACT = " & F2 & " WHERE ID_M_A = " & ID_PARA_EDITAR & ""
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
        cEMPLEADO = Me.TextBox1.Text
        cEMPLEADO = Format(CInt(cEMPLEADO), "0000000000")
        Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + FE + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable3 As New DataTable
        Dim MiDataAdapter3 As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter3 = New SqlDataAdapter("SELECT * FROM [VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE (CODIGO = '" & cEMPLEADO & "') AND (DIA = " & F1 & ") ORDER BY ID", Conexion.CxRRHH)
            MiDataTable3.Clear()
            MiDataAdapter3.Fill(MiDataTable3)
            If MiDataTable3.Rows.Count <> 0 Then
                EXISTE_DIA_EN_SP = True
                ID_PARA_EDITAR = MiDataTable3.Rows(0).Item(7).ToString
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