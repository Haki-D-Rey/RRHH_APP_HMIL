Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Public Class Frm_Dosimetria_Imp
    Private Sub Frm_Dosimetria_Imp_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Call CARGAR_COMBO_DEPARTAMENTOS()
            Me.ComboBox1.Enabled = False
            Me.CheckBox1.Checked = False
            Call CARGAR_COMBO_SERVICIOS()
            Me.ComboBox2.Enabled = False
            Me.CheckBox2.Checked = False
            Me.DateTimePicker1.Value = Mid(Now, 1, 10)
            Me.DateTimePicker2.Value = Mid(Now, 1, 10)
            Me.DateTimePicker1.Enabled = False
            Me.DateTimePicker2.Enabled = False
            Me.CheckBox3.Checked = False
            Me.CheckBox1.Focus()
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Dosimetria_Imp_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Call CARGAR_COMBO_DEPARTAMENTOS()
        Me.ComboBox1.Enabled = False
        Me.CheckBox1.Checked = False
        Call CARGAR_COMBO_SERVICIOS()
        Me.ComboBox2.Enabled = False
        Me.CheckBox2.Checked = False
        Me.DateTimePicker1.Value = Mid(Now, 1, 10)
        Me.DateTimePicker2.Value = Mid(Now, 1, 10)
        Me.DateTimePicker1.Enabled = False
        Me.DateTimePicker2.Enabled = False
        Me.CheckBox3.Checked = False
        Me.CheckBox1.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Call CARGAR_COMBO_DEPARTAMENTOS()
        Me.ComboBox1.Enabled = False
        Me.CheckBox1.Checked = False
        Call CARGAR_COMBO_SERVICIOS()
        Me.ComboBox2.Enabled = False
        Me.CheckBox2.Checked = False
        Me.DateTimePicker1.Value = Mid(Now, 1, 10)
        Me.DateTimePicker2.Value = Mid(Now, 1, 10)
        Me.DateTimePicker1.Enabled = False
        Me.DateTimePicker2.Enabled = False
        Me.CheckBox3.Checked = False
        Me.CheckBox1.Focus()
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_DEPARTAMENTOS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT DISTINCT DEPARTAMENTO FROM [MAESTRO DE HIGIENE DOSIMETRIA] order by DEPARTAMENTO", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DEPARTAMENTO"
                .ValueMember = "DEPARTAMENTO"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_SERVICIOS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT DISTINCT SERVICIO FROM [MAESTRO DE HIGIENE DOSIMETRIA] order by SERVICIO", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox2
                .DataSource = MiDataTable
                .DisplayMember = "SERVICIO"
                .ValueMember = "SERVICIO"
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
    Private Sub CheckBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox1.KeyDown
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
    Private Sub CheckBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox2.KeyDown
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
    Private Sub CheckBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Me.ComboBox1.Enabled = Me.CheckBox1.Checked
    End Sub
    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        Me.ComboBox2.Enabled = Me.CheckBox2.Checked
    End Sub
    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        Me.DateTimePicker1.Enabled = Me.CheckBox3.Checked
        Me.DateTimePicker2.Enabled = Me.CheckBox3.Checked
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If Me.DateTimePicker1.Value > Me.DateTimePicker2.Value Then
            MsgBox("La Fecha Inicial del Periodo Evaluado no puede ser mayor a la Fecha Final", vbInformation, "Mensaje del Sistema")
            Me.DateTimePicker1.Focus()
            Exit Sub
        End If
        Dim D1 As String
        D1 = Mid(Me.DateTimePicker1.Value, 1, 2)
        Dim M1 As String
        M1 = Mid(Me.DateTimePicker1.Value, 4, 2)
        Dim A1 As String
        A1 = Mid(Me.DateTimePicker1.Value, 7, 4)
        Dim D2 As String
        D2 = Mid(Me.DateTimePicker2.Value, 1, 2)
        Dim M2 As String
        M2 = Mid(Me.DateTimePicker2.Value, 4, 2)
        Dim A2 As String
        A2 = Mid(Me.DateTimePicker2.Value, 7, 4)
        If Me.CheckBox3.Checked = True Then
            If SELECCION = "" Then
                SELECCION = "({MAESTRO_DE_HIGIENE_DOSIMETRIA.FECHA_PE_I} >= CDateTime (" & A1 & ", " & M1 & ", " & D1 & ")) AND ({MAESTRO_DE_HIGIENE_DOSIMETRIA.FECHA_PE_F} <= CDateTime (" & A2 & ", " & M2 & ", " & D2 & "))"
            Else
                SELECCION = SELECCION & " AND (({MAESTRO_DE_HIGIENE_DOSIMETRIA.FECHA_PE_I} >= CDateTime (" & A1 & ", " & M1 & ", " & D1 & ")) AND ({MAESTRO_DE_HIGIENE_DOSIMETRIA.FECHA_PE_F} <= CDateTime (" & A2 & ", " & M2 & ", " & D2 & ")))"
            End If
        End If
        If Me.CheckBox1.Checked = True Then
            If Me.ComboBox1.Text <> "" Then
                If SELECCION = "" Then
                    SELECCION = "({MAESTRO_DE_HIGIENE_DOSIMETRIA.DEPARTAMENTO} = '" & Me.ComboBox1.Text & "')"
                Else
                    SELECCION = SELECCION & " AND ({MAESTRO_DE_HIGIENE_DOSIMETRIA.DEPARTAMENTO} = '" & Me.ComboBox1.Text & "')"
                End If
            Else
                MsgBox("Debe seleccionar el Departamento Correctamente", vbInformation, "Mensaje del Sistema")
                Me.ComboBox1.Focus()
                Exit Sub
            End If
        End If
        If Me.CheckBox2.Checked = True Then
            If Me.ComboBox2.Text <> "" Then
                If SELECCION = "" Then
                    SELECCION = "({MAESTRO_DE_HIGIENE_DOSIMETRIA.SERVICIO} = '" & Me.ComboBox2.Text & "')"
                Else
                    SELECCION = SELECCION & " AND ({MAESTRO_DE_HIGIENE_DOSIMETRIA.SERVICIO} = '" & Me.ComboBox2.Text & "')"
                End If
            Else
                MsgBox("Debe seleccionar el Servicio Correctamente", vbInformation, "Mensaje del Sistema")
                Me.ComboBox2.Focus()
                Exit Sub
            End If
        End If

        If Me.RadioButton1.Checked = True Then
            PARAMETRO = 17
        End If
        If Me.RadioButton2.Checked = True Then
            PARAMETRO = 18
        End If
        Frm_Visor_Reportes.CrystalR.ShowExportButton = True
        Frm_Visor_Reportes.CrystalR.ShowPrintButton = True
        Frm_Visor_Reportes.ShowDialog()
        SELECCION = ""
    End Sub
End Class