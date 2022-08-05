Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Editar_Estudios
    Private Sub Frm_Editar_Estudios_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = ""
            Me.TextBox2.Text = ""
            Me.DateTimePicker1.Value = Now
            Me.DateTimePicker1.Checked = False
            Me.TextBox3.Text = ""
            Me.TextBox4.Text = ""
            Me.TextBox5.Text = ""
            Me.CheckBox5.Checked = False
            Me.TextBox6.Text = ""
            Me.TextBox7.Text = ""
            Me.ComboBox1.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Editar_Estudios_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        RemoveHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
        Me.TextBox1.Text = vmesID_M_EST
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_ESTUDIO()
        Me.ComboBox1.Text = vmesD_T_ESTUDIO
        Call CARGAR_COMBO_CAT_DE_ESPECIALIDADES_MEDICAS()
        Me.ComboBox3.Text = vmesD_E_MED
        Me.TextBox2.Text = vmesULT_GRADO_A
        If vmesFECHA_ESTUDIO <> "" Then
            Me.DateTimePicker1.Value = vmesFECHA_ESTUDIO
            Me.DateTimePicker1.Checked = True
        Else
            Me.DateTimePicker1.Checked = False
        End If

        Me.TextBox3.Text = vmesNOMBRE_ESTUDIO
        Me.TextBox4.Text = vmesNOMBRE_C_ESTUDIO
        Me.TextBox5.Text = vmesDIRECC_C_ESTUDIO
        Me.CheckBox5.Checked = vmesESTUDIO_F_PAIS
        Call CARGAR_COMBO_CAT_DE_PAIS()
        Me.ComboBox2.Text = vmesD_PAIS
        Me.TextBox6.Text = vmesTITULO_OBT
        Me.TextBox7.Text = vmesOBSERVACIONES
        Me.ComboBox1.Focus()
        AddHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.DateTimePicker1.Value = Now
        Me.DateTimePicker1.Checked = False
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Me.CheckBox5.Checked = False
        Me.TextBox6.Text = ""
        Me.TextBox7.Text = ""
        Me.ComboBox1.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESPECIALIDADES_MEDICAS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESPECIALIDADES MEDICAS] where ID_T_ESTUDIO = " & Me.ComboBox1.SelectedValue & " order by descripcion", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox3
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_E_MED"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_ESTUDIO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE ESTUDIO] order by ID_T_ESTUDIO", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_T_ESTUDIO"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_PAIS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE PAIS] order by ID_PAIS", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox2
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_PAIS"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
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
    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub CheckBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox5.KeyDown
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
    Private Sub TextBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox6.KeyDown
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "0" Or TextBox1.Text = "" Then
            MsgBox("No se ha generado el Codigo del registro", vbInformation, "Mensaje del Sistema")
            Me.TextBox1.Text = "0"
            Me.Button2.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Tipo de Estudio Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox3.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar la Especialidad Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox3.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox2.Text) = 0 Then
            MsgBox("Debe digitar el Ultimo Año Aprobado", vbInformation, "Mensaje del Sistema")
            Me.TextBox2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox3.Text) = 0 Then
            MsgBox("Debe digitar el Nombre del Estudio Corectamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox3.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox4.Text) = 0 Then
            MsgBox("Debe digitar el Nombre del Centro de Estudio Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox4.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox5.Text) = 0 Then
            MsgBox("Debe digitar la Dirección del Centro de Estudio Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox5.Focus()
            Exit Sub
        End If
        If Me.CheckBox5.Checked = True Then
            If Len(Me.ComboBox2.SelectedValue) = 0 Or Me.ComboBox2.Text = "." Then
                MsgBox("Debe seleccionar el Pais de Estudio Correctamente", vbInformation, "Mensaje del Sistema")
                Me.ComboBox2.Focus()
                Exit Sub
            End If
        End If
        If Me.CheckBox5.Checked = False Then
            If Me.ComboBox2.Text <> "SIN DEFINIR" Then
                Me.CheckBox5.Checked = True
                MsgBox("Se marcará la Ocpion <El Estudio es en el Extranjero?> debido a que se selecciono el Pais de Estudio.", vbInformation, "Mensaje del Sistema")
            End If
        End If
        If Len(Me.ComboBox2.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar correctamente el Pais de Estudio", vbInformation, "Mensaje del Sistema")
            Me.ComboBox2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox6.Text) = 0 Then
            MsgBox("Debe digitar el Nombre del Titulo Obtenido Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox6.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox7.Text) = 0 Then
            Me.TextBox7.Text = ""
        End If
        Call EDITAR_REGISTRO()
        vmesID_M_EST = Me.TextBox1.Text
        CERRAR = False
        Me.Close()
    End Sub
    Private Sub EDITAR_REGISTRO()
        Dim fFEC_F_EST As String
        If Me.DateTimePicker1.Checked = True Then
            fFEC_F_EST = Mid(Me.DateTimePicker1.Value, 1, 10)
        Else
            fFEC_F_EST = "NULL"
        End If
        Dim fechaFEstudio As Object = If(fFEC_F_EST <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFEC_F_EST + "', 105), 23)", "NULL")
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE ESTUDIOS] SET ID_T_ESTUDIO = " & Me.ComboBox1.SelectedValue & ", ID_E_MED = " & Me.ComboBox3.SelectedValue & ", ULT_GRADO_A = '" & Me.TextBox2.Text & "' , FECHA_ESTUDIO = " & fechaFEstudio & ", NOMBRE_ESTUDIO = '" & Me.TextBox3.Text & "', NOMBRE_C_ESTUDIO = '" & Me.TextBox4.Text & "', DIRECC_C_ESTUDIO = '" & Me.TextBox5.Text & "', ESTUDIO_F_PAIS = '" & Me.CheckBox5.Checked & "', ID_PAIS = " & Me.ComboBox2.SelectedValue & ", TITULO_OBT = '" & Me.TextBox6.Text & "', OBSERVACIONES = '" & Me.TextBox7.Text & "'  WHERE ID_M_EST = " & Me.TextBox1.Text & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ComboBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Call CARGAR_COMBO_CAT_DE_ESPECIALIDADES_MEDICAS()
    End Sub
End Class