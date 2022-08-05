Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Datos_Eventos
    Private Sub Frm_Datos_Eventos_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Call Clases.BUSCAR_FECHA_Y_HORA()
        Me.TextBox1.Text = ""
        Me.TextBox6.Text = ""
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_ACCIDENTE()
        Me.DateTimePicker1.Value = Mid(FECHA_SERVIDOR, 1, 10)
        Me.DateTimePicker2.Value = Mid(FECHA_SERVIDOR, 1, 10)
        Me.DateTimePicker4.Value = Mid(FECHA_SERVIDOR, 1, 10)
        Me.DateTimePicker3.Value = Mid(FECHA_SERVIDOR, 1, 10)
        Me.TextBox2.Text = ""
        Me.TextBox9.Text = ""                                           '*
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Call CARGAR_COMBO_CAT_DE_AGENTE_MATERIAL_GENERAL()
        Call CARGAR_COMBO_CAT_DE_AGENTE_MATERIAL_SUB_ESTANDAR()
        Call CARGAR_COMBO_CAT_DE_FORMA_DE_OCURRENCIA()
        Call CARGAR_COMBO_CAT_DE_PARTE_DEL_CUERPO_LESIONADA()
        Me.TextBox19.Text = ""                                          '*
        Call CARGAR_COMBO_CAT_DE_NATURALEZA_DE_LA_LESION()
        Me.TextBox5.Text = ""
        Me.TextBox16.Text = ""                                          '*
        Me.TextBox7.Text = ""
        Me.TextBox8.Text = ""
        Me.TextBox10.Text = ""                                          '*
        Me.DateTimePicker5.Value = Mid(FECHA_SERVIDOR, 1, 10)           '*
        Me.DateTimePicker6.Value = Mid(FECHA_SERVIDOR, 1, 10)           '*
        Me.TextBox11.Text = ""                                          '*
        Me.TextBox12.Text = ""                                          '*
        Me.TextBox13.Text = ""                                          '*
        Me.TextBox14.Text = ""                                          '*
        Me.TextBox15.Text = ""                                          '*
        Me.TextBox18.Text = ""                                          '*
        Me.TextBox17.Text = ""                                          '*
        Me.TextBox6.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox6.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub Frm_Datos_Eventos_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox6.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_NATURALEZA_DE_LA_LESION()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE NATURALEZA DE LA LESION] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox6
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_NL"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_PARTE_DEL_CUERPO_LESIONADA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE PARTE DEL CUERPO LESIONADA] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox5
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_PCL"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_FORMA_DE_OCURRENCIA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE FORMA DE OCURRENCIA] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox4
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_FO"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_AGENTE_MATERIAL_SUB_ESTANDAR()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE AGENTE MATERIAL SUB ESTANDAR] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox3
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_AMSE"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_AGENTE_MATERIAL_GENERAL()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE AGENTE MATERIAL GENERAL] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox2
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_AMG"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_ACCIDENTE()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE ACCIDENTE] order by DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_T_ACCIDENTE"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub TextBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox6.KeyDown
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
    Private Sub DateTimePicker4_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker4.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub DateTimePicker3_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker3.KeyDown
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
    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
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
    Private Sub ComboBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox6.KeyDown
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
    Private Sub TextBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox7.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub GENERAR_CODIGO()
        CODIGO = 0
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            MiDataAdapter = New SqlDataAdapter("SELECT ID_M_EVENTO FROM [MAESTRO DE HIGIENE DATOS DEL EVENTO] ORDER BY ID_M_EVENTO", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                For I = 0 To MiDataTable.Rows.Count - 1
                    CODIGO = MiDataTable.Rows(I).Item(0).ToString
                Next
                CODIGO += 1
            Else
                CODIGO = 1
            End If
            Me.TextBox1.Text = CODIGO
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.TextBox1.Text = "0"
        Call GENERAR_CODIGO()
        If TextBox1.Text = "0" Or TextBox1.Text = "" Then
            MsgBox("No se ha generado el Codigo del registro", vbInformation, "Mensaje del Sistema")
            Me.TextBox1.Text = "0"
            Me.Button2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox6.Text) = 0 Then
            Me.TextBox6.Text = "."
        End If
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Tipo de Accidente Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox2.Text) = 0 Then
            Me.TextBox2.Text = "."
        End If
        If Len(Me.TextBox3.Text) = 0 Then
            Me.TextBox3.Text = "0"
        End If
        If Len(Me.TextBox4.Text) = 0 Then
            Me.TextBox4.Text = "0"
        End If
        If Len(Me.ComboBox2.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Agente Material General Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox2.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox3.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Agente Material Sub Estandar Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox3.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox4.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar la Forma de Ocurrencia Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox4.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox4.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar la Forma de Ocurrencia Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox4.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox5.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar la Parte del Cuerpo Lesionada Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox5.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox6.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar la Naturaleza de la Lesion Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox6.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox5.Text) = 0 Then
            Me.TextBox5.Text = "0"
        End If
        If Len(Me.TextBox7.Text) = 0 Then
            Me.TextBox7.Text = "."
        End If
        If Len(Me.TextBox8.Text) = 0 Then
            Me.TextBox8.Text = "."
        End If
        Call GRABAR_REGISTRO()
        mieID_M_EVENTO = Me.TextBox1.Text
        CERRAR = False
        Me.Close()
    End Sub
    Private Sub GRABAR_REGISTRO()
        Dim fFECHA_ACC As String
        fFECHA_ACC = Mid(Me.DateTimePicker1.Value, 1, 10)
        Dim f1 As Object = If(fFECHA_ACC <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_ACC + "', 105), 23)", "NULL")

        Dim hHORA_ACC As DateTime
        hHORA_ACC = Me.DateTimePicker2.Value
        Dim HORA1 As String = hHORA_ACC.Year & "/" & hHORA_ACC.Month & "/" & hHORA_ACC.Day & " " & hHORA_ACC.Hour & ":" & hHORA_ACC.Minute

        Dim fFECHA_REPORTE As String
        fFECHA_REPORTE = Mid(Me.DateTimePicker4.Value, 1, 10)
        Dim f2 As Object = If(fFECHA_REPORTE <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_REPORTE + "', 105), 23)", "NULL")

        Dim hHORA_REPORTE As DateTime
        hHORA_REPORTE = Me.DateTimePicker3.Value
        Dim HORA2 As String = hHORA_REPORTE.Year & "/" & hHORA_REPORTE.Month & "/" & hHORA_REPORTE.Day & " " & hHORA_REPORTE.Hour & ":" & hHORA_REPORTE.Minute


        Dim fF_II As String
        fF_II = Mid(Me.DateTimePicker5.Value, 1, 10)
        Dim f3 As Object = If(fF_II <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fF_II + "', 105), 23)", "NULL")

        Dim fFCI As String
        fFCI = Mid(Me.DateTimePicker6.Value, 1, 10)
        Dim f4 As Object = If(fFCI <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFCI + "', 105), 23)", "NULL")

        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE HIGIENE DATOS DEL EVENTO] (ID_M_EVENTO, ID_M_IT, ID_M_P, DESCRIPCION_ACC, ID_T_ACCIDENTE, LUGAR_ACC, FECHA_ACC, HORA_ACC, FECHA_REPORTE, HORA_REPORTE, HORAS_LABORADAS, HORAS_PERDIDAS, ID_AMG, ID_AMSE, ID_FO, ID_PCL, ID_NL, DIAS_SUBSIDIO, DECLARACION_T1, DECLARACION_T2, LUGAR_E_ACC, PCL_E, TIEMPO_O, MES, F_II, FCI, ACC_ANT, TURNO_DE_T, HAL, SDT, LQR, ENTREVISTADO_1, ENTREVISTADO_2)" &
            "values (" & CInt(Me.TextBox1.Text) & ", " & h2ID_M_IT & ", " & Frm_Accidentes_Laborales.TextBox36.Text & ", '" & Me.TextBox6.Text & "', " & Me.ComboBox1.SelectedValue & ", '" & Me.TextBox2.Text & "', " & f1 & ", '" & HORA1 & "', " & f2 & ", '" & HORA2 & "', " & CInt(Me.TextBox3.Text) & ", " & CInt(Me.TextBox4.Text) & ", " & Me.ComboBox2.SelectedValue & ", " & Me.ComboBox3.SelectedValue & ", " & Me.ComboBox4.SelectedValue & ", " & Me.ComboBox5.SelectedValue & ", " & Me.ComboBox6.SelectedValue & ", " & CInt(Me.TextBox5.Text) & ", '" & Me.TextBox7.Text & "', '" & Me.TextBox8.Text & "', '" & Me.TextBox9.Text & "', '" & Me.TextBox19.Text & "', '" & Me.TextBox16.Text & "', '" & Me.TextBox10.Text & "', " & f3 & ", " & f4 & ", '" & Me.TextBox11.Text & "', '" & Me.TextBox12.Text & "', '" & Me.TextBox13.Text & "', '" & Me.TextBox14.Text & "', '" & Me.TextBox15.Text & "', '" & Me.TextBox18.Text & "', '" & Me.TextBox17.Text & "')"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub TextBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox9.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox19_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox19.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox16_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox16.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox10.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub DateTimePicker5_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker5.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub DateTimePicker6_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker6.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox11_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox11.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox12_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox12.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox13_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox13.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox14_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox14.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox15_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox15.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox18_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox18.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox17_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox17.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
End Class