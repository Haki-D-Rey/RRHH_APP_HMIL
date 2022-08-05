Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Aspirantes_Upd
    Private Sub Frm_Aspirantes_Upd_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Aspirantes_Upd_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox16.Text = aspID_S
        Me.TextBox1.Text = aspCEDULA
        Me.TextBox2.Text = aspAPELLIDOS
        Me.TextBox3.Text = aspNOMBRES
        Me.DateTimePicker1.Value = aspFECHA_NAC
        Me.TextBox4.Text = aspPERFIL
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_SEXO()
        Me.ComboBox1.Text = aspD_T_SEXO
        Me.TextBox5.Text = aspDIRECCION_1
        Me.TextBox6.Text = aspDIRECCION_2
        Me.TextBox7.Text = aspTELEFONO_1
        Me.TextBox8.Text = aspTELEFONO_2
        Me.TextBox18.Text = aspCORREO_1
        Me.TextBox17.Text = aspCORREO_2
        Me.TextBox9.Text = aspWHATSAPP
        Me.TextBox10.Text = aspPADRE
        Me.TextBox11.Text = aspMADRE
        Me.TextBox12.Text = aspREFERENCIA_ACADEMICA
        Me.TextBox13.Text = aspREFERENCIA_ESTUDIOS
        Me.TextBox14.Text = aspRECOMENDADO_POR
        Me.TextBox15.Text = aspOBSERVACIONES

        If aspFECHA_I <> "" Then
            Me.DateTimePicker2.Value = aspFECHA_I
            Me.DateTimePicker2.Checked = True
        Else
            Me.DateTimePicker2.Checked = False
        End If

        Me.TextBox1.Focus()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_SEXO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE SEXO] ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_T_SEXO"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.TextBox1.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
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
    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
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
    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
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
    Private Sub TextBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox9.KeyDown
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then
            MsgBox("Debe Digitar el No. de Cédula Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox1.Focus()
            Exit Sub
        End If
        Call BUSCAR_CEDULA()
        If EXISTE_CEDULA = True Then
            Me.TextBox6.SelectAll()
            Me.TextBox6.Focus()
            Exit Sub
        End If
        If TextBox16.Text = "0" Or TextBox16.Text = "" Then
            MsgBox("No se ha generado el Codigo del registro", vbInformation, "Mensaje del Sistema")
            Me.TextBox16.Text = "0"
            Me.Button2.Focus()
            Exit Sub
        End If
        If TextBox2.Text = "" Then
            MsgBox("Debe Digitar los Apellidos Aspirante Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox2.Focus()
            Exit Sub
        End If
        If TextBox3.Text = "" Then
            MsgBox("Debe Digitar los Nombres del Aspirante Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox3.Focus()
            Exit Sub
        End If
        If TextBox4.Text = "" Then
            MsgBox("Debe Digitar el Perfil del Aspirante Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox4.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Tipo de Sexo del Aspirante Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If TextBox5.Text = "" Then
            MsgBox("Debe Digitar la Primer Dirección del Aspirante Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox5.Focus()
            Exit Sub
        End If
        If TextBox10.Text = "" Then
            MsgBox("Debe Digitar el Nombre del Padre o Tutor Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox10.Focus()
            Exit Sub
        End If
        If TextBox11.Text = "" Then
            MsgBox("Debe Digitar el Nombre de la Madre o Tutor Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox11.Focus()
            Exit Sub
        End If
        If TextBox12.Text = "" Then
            MsgBox("Debe Digitar la Referencia Academica Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox12.Focus()
            Exit Sub
        End If
        If TextBox13.Text = "" Then
            MsgBox("Debe Digitar la Referencia de Estudios Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox13.Focus()
            Exit Sub
        End If
        'If TextBox14.Text = "" Then
        '    MsgBox("Debe Digitar la Persona\Entidad que recomienda al Aspirante", vbInformation, "Mensaje del Sistema")
        '    Me.TextBox14.Focus()
        '    Exit Sub
        'End If
        Call EDITAR_REGISTRO()
        aspID_S = Me.TextBox16.Text
        CERRAR = False
        Me.Close()
    End Sub
    Private Sub EDITAR_REGISTRO()
        Dim FECHA01 As String = Mid(Me.DateTimePicker1.Value, 1, 10)
        Dim F1 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA01 + "', 105), 23)"

        Dim fFEC_ING As String = Mid(Me.DateTimePicker2.Value, 1, 10)
        If Me.DateTimePicker2.Checked = True Then
            fFEC_ING = Mid(Me.DateTimePicker2.Value, 1, 10)
        Else
            fFEC_ING = "NULL"
        End If
        Dim fechaingreso As Object = If(fFEC_ING <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFEC_ING + "', 105), 23)", "NULL")

        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [SOL_MAESTRO DE SOLICITANTES] SET CEDULA = '" & Trim(Me.TextBox1.Text) & "', APELLIDOS = '" & Trim(Me.TextBox2.Text) & "', NOMBRES = '" & Trim(Me.TextBox3.Text) & "', FECHA_NAC = " & F1 & ", PERFIL = '" & Trim(Me.TextBox4.Text) & "', ID_T_SEXO = " & Me.ComboBox1.SelectedValue & ", DIRECCION_1 = '" & Trim(Me.TextBox5.Text) & "', DIRECCION_2 = '" & Trim(Me.TextBox6.Text) & "', TELEFONO_1 = '" & Trim(Me.TextBox7.Text) & "', TELEFONO_2 = '" & Trim(Me.TextBox8.Text) & "', CORREO_1 = '" & Trim(Me.TextBox18.Text) & "', CORREO_2 = '" & Trim(Me.TextBox17.Text) & "', WHATSAPP = '" & Trim(Me.TextBox9.Text) & "', PADRE = '" & Trim(Me.TextBox10.Text) & "', MADRE = '" & Trim(Me.TextBox11.Text) & "', REFERENCIA_ACADEMICA = '" & Trim(Me.TextBox12.Text) & "', REFERENCIA_ESTUDIOS = '" & Trim(Me.TextBox13.Text) & "', RECOMENDADO_POR = '" & Trim(Me.TextBox14.Text) & "', OBSERVACIONES = '" & Trim(Me.TextBox15.Text) & "', FECHA_I = " & fechaingreso & " WHERE ID_S = " & aspID_S & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim EXISTE_CEDULA As Boolean
    Private Sub BUSCAR_CEDULA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [VISTA SOL_MAESTRO DE SOLICITANTES] WHERE ID_S <> " & Me.TextBox16.Text & " AND CEDULA = '" & Trim(Me.TextBox1.Text) & "'", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                EXISTE_CEDULA = True
                MsgBox("El No. de Cédula que intenta agregar ya existe y pertenece a " & MiDataTable.Rows(0).Item(4).ToString, vbInformation, "Mensaje del Sistema")
            Else
                EXISTE_CEDULA = False
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ADDUPD = False
        Frm_Aspirantes_Cargos.ShowDialog()
    End Sub
    Private Sub DateTimePicker2_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
End Class