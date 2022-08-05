﻿Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Cat_Editar_Estructura_N2
    Private Sub Frm_Cat_Editar_Estructura_N2_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = c3ID_N2
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_1()
        Me.ComboBox2.Text = c3D_N1
        Me.TextBox2.Text = c3ORDN2
        Me.TextBox3.Text = c3DESCRIPCION
        Me.TextBox4.Text = c3SIGLAS
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_ESTRUCTURA()
        Me.ComboBox1.Text = c3D_T_ES
        Me.CheckBox1.Checked = c3MIXTA
        Me.CheckBox2.Checked = c3MILITAR
        Me.CheckBox3.Checked = c3PAME
        Me.TextBox2.Focus()
        Me.TextBox2.SelectAll()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox2.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_1()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTRUCTURA NIVEL 1] order by ORDEN", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox2
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_N1"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_ESTRUCTURA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE ESTRUCTURA] order by ID_T_ES", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_T_ES"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Frm_Cat_Editar_Estructura_N2_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = ""
            Me.TextBox2.Text = ""
            Me.TextBox3.Text = ""
            Me.TextBox4.Text = ""
            Me.TextBox2.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "0" Or TextBox1.Text = "" Then
            MsgBox("No se ha generado el Codigo del registro", vbInformation, "Mensaje del Sistema")
            Me.TextBox1.Text = "0"
            Me.Button2.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox2.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Nivel 1 Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox2.Text) = 0 Then
            MsgBox("Debe digitar el No. de Orden Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox2.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox3.Text) = 0 Then
            MsgBox("Debe digitar la Descripción Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox3.Focus()
            Exit Sub
        End If
        If Len(Me.TextBox4.Text) = 0 Then
            MsgBox("Debe digitar la Sigla Correctamente o escribir (.) si no desea una Sigla para este registro", vbInformation, "Mensaje del Sistema")
            Me.TextBox4.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Tipo Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        Call EDITAR_REGISTRO()
        'Call EDITAR_REGISTRO_EN_N2_EN_MAESTRO_DE_CARGO()
        Call REENUMERAR()
        c3ID_N2 = Me.TextBox1.Text
        Me.TextBox1.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox2.Focus()
        Me.TextBox2.SelectAll()
        CERRAR = False
        Me.Close()
    End Sub
    Dim iINDICE As Integer
    Dim iNORDEN As String
    Private Sub REENUMERAR()
        Dim CONTADOR As Integer
        CONTADOR = 0
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [dbo].[VISTA CAT DE ESTRUCTURA NIVEL 2] ORDER BY [ORDN2]", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                For I = 0 To MiDataTable.Rows.Count - 1
                    CONTADOR = CONTADOR + 1
                    iINDICE = MiDataTable.Rows(I).Item(0).ToString
                    iNORDEN = CONTADOR
                    Call ACTUALIZAR_CAT_DE_ESTRUCTURA_NIVEL_2()
                Next
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ACTUALIZAR_CAT_DE_ESTRUCTURA_NIVEL_2()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [CAT DE ESTRUCTURA NIVEL 2] SET ORDEN = '" & iNORDEN & "' WHERE ID_N2 = " & CInt(iINDICE) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub EDITAR_REGISTRO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [CAT DE ESTRUCTURA NIVEL 2] SET ID_N1 = '" & Me.ComboBox2.SelectedValue & "', ORDEN = '" & Trim(Me.TextBox2.Text) & "', DESCRIPCION = '" & Me.TextBox3.Text & "', SIGLAS = '" & Me.TextBox4.Text & "', ID_T_ES = " & Me.ComboBox1.SelectedValue & ", MIXTA = '" & Me.CheckBox1.Checked & "', MILITAR = '" & Me.CheckBox2.Checked & "', PAME = '" & Me.CheckBox3.Checked & "' WHERE ID_N2 = " & CInt(Me.TextBox1.Text) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    'Private Sub EDITAR_REGISTRO_EN_N2_EN_MAESTRO_DE_CARGO()
    '    If Conexion.CxRRHH.State = ConnectionState.Open Then
    '        Conexion.CBD_RRHH()
    '    End If
    '    Conexion.ABD_RRHH()
    '    Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
    '    Try
    '        MiSqlCommand.CommandText = "UPDATE [MAESTRO DE CARGOS] SET ID_N1 = '" & Me.ComboBox2.SelectedValue & "' WHERE ID_N2 = " & CInt(Me.TextBox1.Text) & ""
    '        MiSqlCommand.ExecuteNonQuery()
    '    Catch ex As Exception
    '        MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
    '    Finally
    '        If Conexion.CxRRHH.State = ConnectionState.Open Then
    '            Conexion.CBD_RRHH()
    '        End If
    '    End Try
    'End Sub
    Private Sub CheckBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox1.KeyDown
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
    Private Sub CheckBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
End Class