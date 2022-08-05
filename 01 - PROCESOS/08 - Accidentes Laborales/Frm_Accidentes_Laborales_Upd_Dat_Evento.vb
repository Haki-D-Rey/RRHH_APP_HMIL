Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Accidentes_Laborales_Upd_Dat_Evento
    Private Sub Frm_Accidentes_Laborales_Upd_Dat_Evento_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox6.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Accidentes_Laborales_Upd_Dat_Evento_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = mieID_M_EVENTO
        Me.TextBox6.Text = mieDESCRIPCION_ACC
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_ACCIDENTE()
        Me.ComboBox1.Text = mieD_T_ACCIDENTE
        Me.DateTimePicker1.Value = mieFECHA_ACC
        Me.DateTimePicker2.Value = mieHORA_ACC
        Me.DateTimePicker4.Value = mieFECHA_REPORTE
        Me.DateTimePicker3.Value = mieHORA_REPORTE
        Me.TextBox2.Text = mieLUGAR_ACC
        Me.TextBox9.Text = mieLUGAR_E_ACC                                    '*
        Me.TextBox3.Text = mieHORAS_LABORADAS
        Me.TextBox4.Text = mieHORAS_PERDIDAS
        Call CARGAR_COMBO_CAT_DE_AGENTE_MATERIAL_GENERAL()
        Me.ComboBox2.Text = mieD_AMG
        Call CARGAR_COMBO_CAT_DE_AGENTE_MATERIAL_SUB_ESTANDAR()
        Me.ComboBox3.Text = mieD_AMSE
        Call CARGAR_COMBO_CAT_DE_FORMA_DE_OCURRENCIA()
        Me.ComboBox4.Text = mieD_FO
        Call CARGAR_COMBO_CAT_DE_PARTE_DEL_CUERPO_LESIONADA()
        Me.TextBox19.Text = miePCL_E                                        '*
        Me.ComboBox5.Text = mieD_PCL
        Call CARGAR_COMBO_CAT_DE_NATURALEZA_DE_LA_LESION()
        Me.ComboBox6.Text = mieD_NL
        Me.TextBox5.Text = mieDIAS_SUBSIDIO
        Me.TextBox16.Text = mieTIEMPO_O                                     '*
        Me.TextBox7.Text = mieDECLARACION_T1
        Me.TextBox8.Text = mieDECLARACION_T2
        Me.TextBox10.Text = mieMES                                          '*
        Me.DateTimePicker5.Value = Mid(mieF_II, 1, 10)                      '*
        Me.DateTimePicker6.Value = Mid(mieFCI, 1, 10)                       '*
        Me.TextBox11.Text = mieACC_ANT                                      '*
        Me.TextBox12.Text = mieTURNO_DE_T                                   '*
        Me.TextBox13.Text = mieHAL                                          '*
        Me.TextBox14.Text = mieSDT                                          '*
        Me.TextBox15.Text = mieLQR                                          '*
        Me.TextBox18.Text = mieENTREVISTADO_1                               '*
        Me.TextBox17.Text = mieENTREVISTADO_2                               '*
        Me.TextBox6.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.TextBox6.Focus()
        CERRAR = True
        Me.Close()
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
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
        Call EDITAR_REGISTRO()
        mieID_M_EVENTO = Me.TextBox1.Text
        CERRAR = False
        Me.Close()
    End Sub
    Private Sub EDITAR_REGISTRO()
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
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE HIGIENE DATOS DEL EVENTO] SET DESCRIPCION_ACC = '" & Me.TextBox6.Text & "', ID_T_ACCIDENTE = " & Me.ComboBox1.SelectedValue & ", LUGAR_ACC= '" & Me.TextBox2.Text & "', FECHA_ACC = " & f1 & ", HORA_ACC= '" & HORA1 & "', FECHA_REPORTE = " & f2 & ", HORA_REPORTE = '" & HORA2 & "', HORAS_LABORADAS = " & CInt(Me.TextBox3.Text) & ", HORAS_PERDIDAS = " & CInt(Me.TextBox4.Text) & ", ID_AMG = " & Me.ComboBox2.SelectedValue & ", ID_AMSE = " & Me.ComboBox3.SelectedValue & ", ID_FO = " & Me.ComboBox4.SelectedValue & ", ID_PCL = " & Me.ComboBox5.SelectedValue & ", ID_NL = " & Me.ComboBox6.SelectedValue & ", DIAS_SUBSIDIO = " & CInt(Me.TextBox5.Text) & ", DECLARACION_T1 = '" & Me.TextBox7.Text & "', DECLARACION_T2 = '" & Me.TextBox8.Text & "', LUGAR_E_ACC = '" & Me.TextBox9.Text & "', PCL_E = '" & Me.TextBox19.Text & "', TIEMPO_O = '" & Me.TextBox16.Text & "', MES = '" & Me.TextBox10.Text & "', F_II = " & f3 & ", FCI = " & f4 & ", ACC_ANT = '" & Me.TextBox11.Text & "', TURNO_DE_T = '" & Me.TextBox12.Text & "', HAL = '" & Me.TextBox13.Text & "', SDT = '" & Me.TextBox14.Text & "', LQR = '" & Me.TextBox15.Text & "', ENTREVISTADO_1 = '" & Me.TextBox18.Text & "', ENTREVISTADO_2 = '" & Me.TextBox17.Text & "' WHERE ID_M_EVENTO = " & CInt(Me.TextBox1.Text) & ""
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