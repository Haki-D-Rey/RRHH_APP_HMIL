Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Editar_CyE
    Private Sub Frm_Editar_CyE_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = "0"
            Me.TextBox3.Text = ""
            Me.TextBox2.Text = ""
            Me.ComboBox1.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Editar_CyE_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = mcyeID_M_CYE
        Call CARGAR_COMBO_CAT_DE_CONTRATOS()
        Me.ComboBox1.Text = mcyeDESCRIPCION
        If mcyeFECHA_I <> "" Then
            Me.DateTimePicker1.Value = mcyeFECHA_I
            Me.DateTimePicker1.Checked = True
        Else
            Me.DateTimePicker1.Checked = False
        End If
        If mcyeFECHA_F <> "" Then
            Me.DateTimePicker2.Value = mcyeFECHA_F
            Me.DateTimePicker2.Checked = True
        Else
            Me.DateTimePicker2.Checked = False
        End If
        Me.CheckBox1.Checked = mcyeFIRMADO

        If mcyeFECHA_EVA <> "" Then
            Me.DateTimePicker3.Value = mcyeFECHA_EVA
            Me.DateTimePicker3.Checked = True
        Else
            Me.DateTimePicker3.Checked = False
        End If

        If mcyePROMEDIO_NOTA <> "" Then
            Me.TextBox3.Text = mcyePROMEDIO_NOTA
        Else
            Me.TextBox3.Text = ""
        End If
        Call CARGAR_COMBO_CAT_DE_CLASIFICACION_DE_EVALUACIONES()

        If mcyeD_C_E <> "" Then
            Me.ComboBox2.Text = mcyeD_C_E
        Else
            Me.ComboBox2.Text = ""
        End If
        Me.TextBox2.Text = mcyeOBSERVACIONES
        Me.ComboBox1.Focus()
    End Sub
    Private Sub Button7_Click(sender As Object, e As System.EventArgs) Handles Button7.Click
        Me.TextBox1.Text = "0"
        Me.TextBox3.Text = ""
        Me.TextBox2.Text = ""
        Me.ComboBox1.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_CLASIFICACION_DE_EVALUACIONES()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE CLASIFICACION DE EVALUACIONES] order by ID_C_E", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox2
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_C_E"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_CONTRATOS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE CONTRATOS] order by ID_CONT", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_CONT"
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
    Private Sub CheckBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox1.KeyDown
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "0" Or TextBox1.Text = "" Then
            MsgBox("No se ha generado el Codigo del registro", vbInformation, "Mensaje del Sistema")
            Me.TextBox1.Text = "0"
            Me.Button7.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Tipo Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox2.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar la Clasificacion Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox2.Text) = 0 Then
            Me.TextBox2.Text = ""
        End If
        If Me.TextBox3.Text = "" Then
            Me.TextBox3.Text = "0"
        End If
        If Me.ComboBox1.Text = "1° CONTRATO" Or Me.ComboBox1.Text = "2° CONTRATO" Or Me.ComboBox1.Text = "CONTRATO INDEFINIDO" Then
            Me.TextBox3.Text = ""
            Me.ComboBox2.Text = "SIN DEFINIR"
            GoTo SALTO
        End If
        If Not IsNumeric(Me.TextBox3.Text) Then
            Me.TextBox3.Text = "0"
        End If
        If IsNumeric(CInt(Me.TextBox3.Text)) Then
            Dim vNOTA As Double
            vNOTA = CDbl(Me.TextBox3.Text)
            If vNOTA < 60 Then
                Me.ComboBox2.Text = "DEFICIENTE"
            End If
            If vNOTA > 59 And vNOTA < 69 Then
                Me.ComboBox2.Text = "REGULAR"
            End If
            If vNOTA > 68 And vNOTA < 79 Then
                Me.ComboBox2.Text = "BUENO"
            End If
            If vNOTA > 78 And vNOTA < 89 Then
                Me.ComboBox2.Text = "MUY BUENO"
            End If
            If vNOTA > 88 And vNOTA < 101 Then
                Me.ComboBox2.Text = "EXCELENTE"
            End If
        Else
            MsgBox("El dato ingresado no es númerico o no es váildo", vbInformation, "Mensaje del Sistema")
            Me.TextBox3.SelectAll()
            Me.TextBox3.Focus()
            Exit Sub
        End If
SALTO:
        Call EDITAR_REGISTRO()
        mcyeID_M_CYE = Me.TextBox1.Text
        Me.TextBox1.Text = "0"
        Me.TextBox2.Text = ""
        Me.ComboBox1.Focus()
        CERRAR = False
        Me.Close()
    End Sub
    Private Sub EDITAR_REGISTRO()
        Dim fFECHA_I As String
        If Me.DateTimePicker1.Checked = True Then
            fFECHA_I = Mid(Me.DateTimePicker1.Value, 1, 10)
        Else
            fFECHA_I = "NULL"
        End If
        Dim f1 As Object = If(fFECHA_I <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_I + "', 105), 23)", "NULL")

        Dim fFECHA_F As String
        If Me.DateTimePicker2.Checked = True Then
            fFECHA_F = Mid(Me.DateTimePicker2.Value, 1, 10)
        Else
            fFECHA_F = "NULL"
        End If
        Dim f2 As Object = If(fFECHA_F <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_F + "', 105), 23)", "NULL")

        Dim fFECHA_E As String
        If Me.DateTimePicker3.Checked = True Then
            fFECHA_E = Mid(Me.DateTimePicker3.Value, 1, 10)
        Else
            fFECHA_E = "NULL"
        End If
        Dim f3 As Object = If(fFECHA_E <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_E + "', 105), 23)", "NULL")


        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE CONTRATOS Y EVALUACIONES] SET ID_CONT = " & Me.ComboBox1.SelectedValue & ", FECHA_I = " & f1 & ", FECHA_F = " & f2 & ", FIRMADO = '" & Me.CheckBox1.Checked & "', FECHA_EVA = " & f3 & ", PROMEDIO_NOTA = '" & Me.TextBox3.Text & "', ID_C_E = " & Me.ComboBox2.SelectedValue & ", OBSERVACIONES = '" & Me.TextBox2.Text & "' WHERE ID_M_CYE = " & Me.TextBox1.Text & ""
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
            If Me.TextBox3.Text = "" Then
                Me.TextBox3.Text = "0"
            End If
            If Not IsNumeric(Me.TextBox3.Text) Then
                Me.TextBox3.Text = "0"
            End If
            If IsNumeric(CInt(Me.TextBox3.Text)) Then
                Dim vNOTA As Double
                vNOTA = CDbl(Me.TextBox3.Text)
                If vNOTA < 60 Then
                    Me.ComboBox2.Text = "DEFICIENTE"
                End If
                If vNOTA > 59 And vNOTA < 69 Then
                    Me.ComboBox2.Text = "REGULAR"
                End If
                If vNOTA > 68 And vNOTA < 79 Then
                    Me.ComboBox2.Text = "BUENO"
                End If
                If vNOTA > 78 And vNOTA < 89 Then
                    Me.ComboBox2.Text = "MUY BUENO"
                End If
                If vNOTA > 88 And vNOTA < 101 Then
                    Me.ComboBox2.Text = "EXCELENTE"
                End If
            Else
                MsgBox("El dato ingresado no es númerico o no es váildo", vbInformation, "Mensaje del Sistema")
                Me.TextBox3.SelectAll()
                Me.TextBox3.Focus()
                Exit Sub
            End If
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
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
End Class