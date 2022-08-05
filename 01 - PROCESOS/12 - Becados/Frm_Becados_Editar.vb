Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Becados_Editar
    Private Sub Frm_Becados_Editar_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox4.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Becados_Editar_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Call INICIAR_OBJETOS()
        Me.TextBox4.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox4.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub INICIAR_OBJETOS()
        Me.TextBox1.Text = mbID_B
        Me.TextBox2.Text = mbNOMBRES
        Me.TextBox3.Text = mbAPELLIDOS
        Me.TextBox4.Text = mbCODIGO
        Me.TextBox5.Text = mbN_CEDULA
        Me.TextBox6.Text = mbN_MINSA
        Me.TextBox7.Text = mbD_T_SEXO
        Call CARGAR_COMBO_CAT_DE_PAIS()
        Me.ComboBox1.Text = mbD_PAIS
        Me.TextBox18.Text = mbNOMBRE_ESTUDIO
        Me.TextBox8.Text = mbCENTRO_ESTUDIO
        Me.DateTimePicker2.Value = mbFEC_INI
        Me.DateTimePicker3.Value = mbFEC_FIN
        Me.TextBox9.Text = mbTELEFONO
        Me.TextBox10.Text = mbNO_CUENTA
        Me.TextBox11.Text = mbCORREO
        Me.TextBox12.Text = mbMATRICULA
        Me.TextBox13.Text = mbINSCRIPCION
        Me.TextBox14.Text = mbALIMENTACION
        Me.TextBox15.Text = mbHOSPEDAJE
        Me.TextBox16.Text = mbPAGO_TITULO
        Me.TextBox20.Text = mbOTROS
        Me.TextBox17.Text = mbOBSERVACIONES
        Me.TextBox19.Text = mbID_M_P
        Me.TextBox4.Focus()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_PAIS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE PAIS] ORDER BY ID_PAIS", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
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
    Dim cEMPLEADO As String
    Dim DATO As Boolean
    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            cEMPLEADO = CInt(Val(TextBox4.Text))
            TextBox4.Text = Format(CInt(cEMPLEADO), "0000000000")
            If Len(Me.TextBox4.Text) = 0 Or (Me.TextBox4.Text = "0000000000") Then
                MsgBox("El código que se digitó no es válido", vbInformation, "Mensaje del Sistema")
                Me.TextBox1.Focus()
                Exit Sub
            Else
                Call BUSCAR_DATO()
                If DATO = True Then
                    e.Handled = True
                    SendKeys.Send("{TAB}")
                Else
                    MsgBox("El código que se digitó no es válido", vbInformation, "Mensaje del Sistema")
                    Me.TextBox4.SelectAll()
                    Me.TextBox4.Focus()
                    Exit Sub
                End If
            End If
        End If
    End Sub
    Private Sub BUSCAR_DATO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT ID_M_P, CODIGO, APELLIDOS, NOMBRES, N_CEDULA, N_MINSA, D_T_SEXO, ID_ESTADO_P FROM [dbo].[VISTA MAESTRO DE PERSONAS] WHERE CODIGO = '" & Me.TextBox4.Text & "' AND ID_ESTADO_P = 1", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                Me.TextBox19.Text = MiDataTable.Rows(0).Item(0).ToString
                Me.TextBox2.Text = MiDataTable.Rows(0).Item(2).ToString
                Me.TextBox3.Text = MiDataTable.Rows(0).Item(3).ToString
                Me.TextBox5.Text = MiDataTable.Rows(0).Item(4).ToString
                Me.TextBox6.Text = MiDataTable.Rows(0).Item(5).ToString
                Me.TextBox7.Text = MiDataTable.Rows(0).Item(6).ToString
                DATO = True
            Else
                Me.TextBox1.Text = ""
                Me.TextBox2.Text = ""
                Me.TextBox3.Text = ""
                Me.TextBox5.Text = ""
                Me.TextBox6.Text = ""
                Me.TextBox7.Text = ""
                DATO = False
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
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
    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
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
    Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyDown
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
    Private Sub DateTimePicker3_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker3.KeyDown
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
    Private Sub TextBox16_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox16.KeyDown
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
        If TextBox19.Text = "0" Or TextBox19.Text = "" Then
            MsgBox("No se encuentra el Codigo del Empleado", vbInformation, "Mensaje del Sistema")
            Me.TextBox19.Text = "0"
            Me.Button2.Focus()
            Exit Sub
        End If
        If TextBox1.Text = "0" Or TextBox1.Text = "" Then
            MsgBox("No se ha generado el Codigo del registro", vbInformation, "Mensaje del Sistema")
            Me.TextBox1.Text = "0"
            Me.Button2.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Pais de Estudio Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If TextBox18.Text = "" Then
            MsgBox("Debe Digitar el Nombre del Estudio Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox18.Focus()
            Exit Sub
        End If
        If TextBox8.Text = "" Then
            MsgBox("Debe Digitar el Centro de Estudio Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox8.Focus()
            Exit Sub
        End If
        If TextBox9.Text = "" Then
            TextBox9.Text = "0"
        End If
        If TextBox10.Text = "" Then
            TextBox10.Text = "."
        End If
        If TextBox11.Text = "" Then
            TextBox11.Text = "."
        End If
        If TextBox12.Text = "" Then
            TextBox12.Text = "0"
        End If
        If TextBox13.Text = "" Then
            TextBox13.Text = "0"
        End If
        If TextBox14.Text = "" Then
            TextBox14.Text = "0"
        End If
        If TextBox15.Text = "" Then
            TextBox15.Text = "0"
        End If
        If TextBox16.Text = "" Then
            TextBox16.Text = "0"
        End If
        If TextBox20.Text = "" Then
            TextBox20.Text = "0"
        End If
        If TextBox17.Text = "" Then
            TextBox17.Text = "."
        End If
        TOT = CDbl(Me.TextBox12.Text) + CDbl(Me.TextBox13.Text) + CDbl(Me.TextBox14.Text) + CDbl(Me.TextBox15.Text) + CDbl(Me.TextBox16.Text) + CDbl(Me.TextBox20.Text)
        Call Me.EDITAR_REGISTRO()
        CERRAR = False
        mbID_B = Me.TextBox1.Text
        Call INICIAR_OBJETOS()
        Me.TextBox4.Focus()
        Me.Close()
    End Sub
    Dim TOT As String
    Private Sub EDITAR_REGISTRO()
        Dim fFEC_1 As String = Me.DateTimePicker2.Text
        Dim fFEC_2 As String = Me.DateTimePicker3.Text
        Dim F1 = "Convert(VARCHAR(10), Convert(Date, '" + fFEC_1 + "', 105), 23)"
        Dim F2 = "Convert(VARCHAR(10), Convert(Date, '" + fFEC_2 + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()

        Dim SQL As String
        Try
            SQL = "UPDATE [MAESTRO DE BECADOS] SET ID_M_P = " & Me.TextBox19.Text & ", NOMBRE_ESTUDIO = '" & Me.TextBox18.Text & "', ID_PAIS = " & Me.ComboBox1.SelectedValue & ", CENTRO_ESTUDIO = '" & Me.TextBox8.Text & "', FEC_INI  = " & F1 & ", FEC_FIN = " & F2 & ", TELEFONO = '" & Me.TextBox9.Text & "', NO_CUENTA = '" & Me.TextBox10.Text & "', CORREO = '" & Me.TextBox11.Text & "', MATRICULA = '" & Me.TextBox12.Text & "', INSCRIPCION = '" & Me.TextBox13.Text & "', ALIMENTACION = '" & Me.TextBox14.Text & "', HOSPEDAJE = '" & Me.TextBox15.Text & "', PAGO_TITULO = '" & Me.TextBox16.Text & "', OTROS = '" & Me.TextBox20.Text & "', TOTAL = '" & TOT & "', OBSERVACIONES = '" & Me.TextBox17.Text & "' WHERE ID_B = " & Me.TextBox1.Text & ""
            Dim cmd As New SqlCommand(SQL, CxRRHH)
            cmd.ExecuteNonQuery()
            cmd = Nothing
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub TextBox20_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox20.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
End Class