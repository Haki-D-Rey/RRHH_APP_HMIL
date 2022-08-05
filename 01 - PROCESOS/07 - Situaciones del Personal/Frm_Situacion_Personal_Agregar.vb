Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Situacion_Personal_Agregar
    Private Sub Frm_Situacion_Personal_Agregar_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = "0"
            Call Clases.BUSCAR_FECHA_Y_HORA()
            Me.DateTimePicker10.Value = Mid(FECHA_SERVIDOR, 1, 10)
            Me.DateTimePicker20.Value = Mid(FECHA_SERVIDOR, 1, 10)
            Me.DateTimePicker30.Value = Mid(FECHA_SERVIDOR, 1, 10)
            Me.TextBox2.Text = ""
            Me.ComboBox2.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Situacion_Personal_Agregar_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RemoveHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
        Me.TextBox1.Text = "0"
        Call Clases.BUSCAR_FECHA_Y_HORA()
        Me.DateTimePicker10.Value = Mid(FECHA_SERVIDOR, 1, 10)
        Me.DateTimePicker20.Value = Mid(FECHA_SERVIDOR, 1, 10)
        Me.DateTimePicker30.Value = Mid(FECHA_SERVIDOR, 1, 10)
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_SITUACION_DE_PERSONAL()
        Call CARGAR_COMBO_CAT_DE_SITUACION_DE_PERSONAL()
        Me.TextBox3.Text = "0"
        Me.TextBox2.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Me.TextBox6.Text = ""
        AddHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
        Me.ComboBox2.SelectAll()
        Me.ComboBox2.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox1.Text = "0"
        Call Clases.BUSCAR_FECHA_Y_HORA()
        Me.DateTimePicker10.Value = Mid(FECHA_SERVIDOR, 1, 10)
        Me.DateTimePicker20.Value = Mid(FECHA_SERVIDOR, 1, 10)
        Me.DateTimePicker30.Value = Mid(FECHA_SERVIDOR, 1, 10)
        Me.TextBox2.Text = ""
        Me.ComboBox2.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_SITUACION_DE_PERSONAL()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE SITUACION DE PERSONAL] order by ID_T_SIT_P", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_T_SIT_P"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_SITUACION_DE_PERSONAL()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE SITUACION DE PERSONAL] WHERE (ID_T_SIT_P = " & Me.ComboBox1.SelectedValue & ") AND (ID_SIT_P <> 31) order by ID_T_SIT_P, DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox2
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
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Call CARGAR_COMBO_CAT_DE_SITUACION_DE_PERSONAL()
    End Sub
    Private Sub DateTimePicker1_KeyDown(sender As Object, e As KeyEventArgs)
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
    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub DateTimePicker10_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker10.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub DateTimePicker30_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker30.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Dim FA As String
    Dim DIA_LIBRE_1, DIA_LIBRE_2 As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.ComboBox2.Text = "AJUSTE MENSUAL (+)" Or Me.ComboBox2.Text = "AJUSTE MENSUAL (-)" Then
            MsgBox("El Tipo de Situación no es válido", vbInformation, "Mensaje del Sistema")
            Me.ComboBox2.Focus()
            Exit Sub
        End If
        'If TextBox1.Text = "0" Or TextBox1.Text = "" Then
        '    MsgBox("No se ha generado el Codigo del registro", vbInformation, "Mensaje del Sistema")
        '    Me.TextBox1.Text = "0"
        '    Me.Button2.Focus()
        '    Exit Sub
        'End If
        If Me.ComboBox2.Text = "SUBSIDIO" Then
            If Len(Me.TextBox4.Text) < 4 Then
                MsgBox("Debe ingresar los Datos Generales del Subsidio Correctamente", vbInformation, "Mensaje del Sistema")
                Me.TextBox4.SelectAll()
                Me.TextBox4.Focus()
                Exit Sub
            End If
            If Len(Me.TextBox6.Text) < 4 Then
                MsgBox("Debe ingresar el Diagnostico Medico del Subsidio Correctamente", vbInformation, "Mensaje del Sistema")
                Me.TextBox6.SelectAll()
                Me.TextBox6.Focus()
                Exit Sub
            End If
        End If
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Tipo de Situacion Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox2.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar la Situacion Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox2.Focus()
            Exit Sub
        End If
        Call RECORRER_DIAS()
        msID_M_A = Me.TextBox1.Text
        CERRAR = False
        Me.ComboBox2.Focus()
        Me.Close()
    End Sub
    Dim ANTxDIA As Double
    Private Sub RECORRER_DIAS()
        vDET1 = " "
        vDET2 = " "
        vDET3 = " "
        Me.DateTimePicker20.Value = Me.DateTimePicker10.Value
        'Me.DateTimePicker20.Value = Mid(Now, 1, 10)            'ESTA ORIGINAL
        'Me.DateTimePicker2.Value = Mid(Now, 1, 10)             'ESTA ORIGINAL

        Me.Cursor = Cursors.WaitCursor
        Application.DoEvents()
        'For I2 As Integer = 1 To DateDiff(DateInterval.Day, Me.DateTimePicker10.Value, Me.DateTimePicker30.Value.AddDays(1))
        For I2 As Integer = 1 To DateDiff(DateInterval.Day, Me.DateTimePicker10.Value, Me.DateTimePicker30.Value.AddDays(1))
            Me.DateTimePicker30.Refresh()
            Call BUSCAR_SI_EXISTE_DIA_EN_MAESTRO_DE_SITUACION_DE_PERSONAL()
            If EXISTE_DIA_EN_MAESTRO_DE_SITUACION_DE_PERSONAL = True Then
                GoTo LEE_OTRO
            Else
                Call AGREGAR_REGISTRO()
            End If
LEE_OTRO:
            Me.DateTimePicker20.Value = Me.DateTimePicker20.Value.AddDays(1)
            Me.DateTimePicker20.Refresh()
        Next
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub AGREGAR_REGISTRO()
        Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + Me.DateTimePicker20.Value + "', 105), 23)"
        Call Clases.BUSCAR_FECHA_Y_HORA()
        Dim F2 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(FECHA_SERVIDOR, 1, 10) + "', 105), 23)"
        Dim CD As Integer = DateTime.DaysInMonth(Me.DateTimePicker30.Value.Year, Me.DateTimePicker30.Value.Month)
        ANTxDIA = 0.0833333333333333
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE SITUACION DE PERSONAL] (ID_M_P, DIA, ID_T_SIT_P, ID_SIT_P, VALOR_SIT, VALOR_DIA, OBSERVACIONES, DETALLE_1, DETALLE_2, DETALLE_3, USUARIO_ACT, FECHA_ACT)" &
                              "values (" & id_EMPLEADO & ", " & F1 & ", " & Me.ComboBox1.SelectedValue & ", " & Me.ComboBox2.SelectedValue & ", '" & CDbl(Me.TextBox3.Text) & "', '" & CDbl(ANTxDIA) & "','" & Trim(Me.TextBox2.Text) & "', '" & vDET1 & "', '" & vDET2 & "', '" & vDET3 & "', '" & isuUSUARIO & "', " & F2 & ")"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim EXISTE_DIA_EN_MAESTRO_DE_SITUACION_DE_PERSONAL As Boolean
    Private Sub BUSCAR_SI_EXISTE_DIA_EN_MAESTRO_DE_SITUACION_DE_PERSONAL()
        Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(Me.DateTimePicker20.Value, 1, 10) + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable3 As New DataTable
        Dim MiDataAdapter3 As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter3 = New SqlDataAdapter("SELECT CODIGO, DIA FROM [VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE (CODIGO = '" & Frm_Situacion_Personal.TextBox2.Text & "') AND (DIA = " & F1 & ") ORDER BY  CODIGO, DIA", Conexion.CxRRHH)
            MiDataTable3.Clear()
            MiDataAdapter3.Fill(MiDataTable3)
            If MiDataTable3.Rows.Count <> 0 Then
                EXISTE_DIA_EN_MAESTRO_DE_SITUACION_DE_PERSONAL = True
                vDET1 = " "
                vDET2 = " "
                vDET3 = " "
            Else
                EXISTE_DIA_EN_MAESTRO_DE_SITUACION_DE_PERSONAL = False
                vDET1 = " "
                vDET2 = " "
                vDET3 = " "
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