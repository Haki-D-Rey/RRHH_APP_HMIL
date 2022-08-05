Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Situacion_Personal_Editar
    Private Sub Frm_Situacion_Personal_Editar_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = "0"
            Me.DateTimePicker1.Value = Now
            Me.TextBox2.Text = ""
            Me.ComboBox2.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Situacion_Personal_Editar_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        RemoveHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
        Me.TextBox1.Text = msID_M_A
        Me.DateTimePicker1.Value = msDIA
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_SITUACION_DE_PERSONAL()
        Me.ComboBox1.Text = msD_T_SIT_P
        Call CARGAR_COMBO_CAT_DE_SITUACION_DE_PERSONAL()
        Me.ComboBox2.Text = msD_SIT_P
        Me.TextBox3.Text = msVALOR_SIT
        Me.TextBox2.Text = msOBSERVACIONES
        Me.TextBox4.Text = msDETALLE_1
        Me.TextBox5.Text = msDETALLE_2
        Me.TextBox6.Text = msDETALLE_3
        AddHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
        Me.ComboBox2.SelectAll()
        Me.ComboBox2.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox1.Text = "0"
        Me.DateTimePicker1.Value = Now
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
    Private Sub DateTimePicker1_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker1.KeyDown
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
    Dim FA As String
    Dim DIA_LIBRE_1, DIA_LIBRE_2 As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.ComboBox2.Text = "AJUSTE MENSUAL (+)" Or Me.ComboBox2.Text = "AJUSTE MENSUAL (-)" Then
            MsgBox("El Tipo de Situación no es válido", vbInformation, "Mensaje del Sistema")
            Me.ComboBox2.Focus()
            Exit Sub
        End If
        If TextBox1.Text = "0" Or TextBox1.Text = "" Then
            MsgBox("No se ha generado el Codigo del registro", vbInformation, "Mensaje del Sistema")
            Me.TextBox1.Text = "0"
            Me.Button2.Focus()
            Exit Sub
        End If
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
        If Frm_Situacion_Personal.CheckBox1.Checked = False Then
            Call EDITAR_REGISTRO()
        Else
            Call RECORRER_DGV()
        End If
        msID_M_A = Me.TextBox1.Text
        CERRAR = False
        Me.ComboBox2.Focus()
        Me.Close()
    End Sub
    Dim ANTxDIA As Double
    Private Sub AGREGAR_REGISTRO_PT12()
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
                              "values (" & id_EMPLEADO & ", " & F1 & ", 2, 29, '" & CDbl(Me.TextBox3.Text) & "', '" & CDbl(ANTxDIA) & "','.', '" & Me.TextBox4.Text & "', '" & Me.TextBox5.Text & "', '" & Me.TextBox6.Text & "', '" & isuUSUARIO & "', " & F2 & ")"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub AGREGAR_REGISTRO_PT24()
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
                              "values (" & id_EMPLEADO & ", " & F1 & ", 2, 29, '" & CDbl(Me.TextBox3.Text) & "', '" & CDbl(ANTxDIA) & "','.', '" & Me.TextBox4.Text & "', '" & Me.TextBox5.Text & "', '" & Me.TextBox6.Text & "', '" & isuUSUARIO & "', " & F2 & ")"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim EXISTE_DIA_EN_SP As Boolean
    Dim FE As String
    Dim ID_PARA_EDITAR As Integer
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
            MiDataAdapter3 = New SqlDataAdapter("SELECT * FROM [VISTA MAESTRO DE SITUACION DE PERSONAL] WHERE (CODIGO = '" & cEMPLEADO & "') AND (DIA = " & F1 & ") ORDER BY ID", Conexion.CxRRHH)
            MiDataTable3.Clear()
            MiDataAdapter3.Fill(MiDataTable3)
            If MiDataTable3.Rows.Count <> 0 Then
                EXISTE_DIA_EN_SP = True
                ID_PARA_EDITAR = MiDataTable3.Rows(0).Item(4).ToString
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
    Private Sub EDITAR_REGISTRO()
        If Me.ComboBox2.Text = "AUSENCIA INJUSTIFICADA" Or Me.ComboBox2.Text = "VACACIONES" Or Me.ComboBox2.Text = "OMISION DE MARCA" Then
            If Val(Me.TextBox3.Text) > 0 Then
                Me.TextBox3.Text = Val(Me.TextBox3.Text) * (-1)
            End If
        End If
        Call Clases.BUSCAR_FECHA_Y_HORA()
        Dim F2 As String = "Convert(VARCHAR(10), Convert(Date, '" + Mid(FECHA_SERVIDOR, 1, 10) + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE SITUACION DE PERSONAL] SET ID_T_SIT_P = " & Me.ComboBox1.SelectedValue & ", ID_SIT_P = " & Me.ComboBox2.SelectedValue & ", VALOR_SIT = '" & CDbl(Me.TextBox3.Text) & "', OBSERVACIONES = '" & Me.TextBox2.Text & "', DETALLE_1='" & Me.TextBox4.Text & "', DETALLE_2='" & Me.TextBox5.Text & "', DETALLE_3= '" & Me.TextBox6.Text & "', USUARIO_ACT = '" & isuUSUARIO & "', FECHA_ACT = " & F2 & "  WHERE ID_M_A = " & msID_M_A & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        '3.17.20
        If Me.ComboBox2.Text = "AUSENCIA INJUSTIFICADA" Or Me.ComboBox2.Text = "VACACIONES" Or Me.ComboBox2.Text = "OMISION DE MARCA" Then
            Me.TextBox3.Text = -1
        Else
            Me.TextBox3.Text = 0
        End If
        If Me.ComboBox2.Text = "DIA ACREDITADO" Then
            Me.TextBox3.Text = 1
        End If
    End Sub
    Private Sub RECORRER_DGV()
        Dim Z As Integer
        Z = 0
        If Frm_Situacion_Personal.DGV.RowCount > 0 Then
            For Z = 0 To Frm_Situacion_Personal.DGV.RowCount - 1
                If Frm_Situacion_Personal.DGV.Rows(Z).Selected = True Then
                    msID_M_A = Frm_Situacion_Personal.DGV.Item(7, Z).Value
                    Call EDITAR_REGISTRO()
                End If
            Next
        End If
    End Sub
    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox6.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
End Class