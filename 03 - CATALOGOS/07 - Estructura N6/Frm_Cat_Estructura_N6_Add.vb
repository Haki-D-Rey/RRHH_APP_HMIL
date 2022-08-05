Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Cat_Estructura_N6_Add
    Private Sub Frm_Cat_Estructura_N6_Add_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
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
    Private Sub Frm_Cat_Estructura_N6_Add_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = ""
        RemoveHandler Me.ComboBox2.SelectedIndexChanged, AddressOf ComboBox2_SelectedIndexChanged
        RemoveHandler Me.ComboBox3.SelectedIndexChanged, AddressOf ComboBox3_SelectedIndexChanged
        RemoveHandler Me.ComboBox4.SelectedIndexChanged, AddressOf ComboBox4_SelectedIndexChanged
        RemoveHandler Me.ComboBox5.SelectedIndexChanged, AddressOf ComboBox5_SelectedIndexChanged
        RemoveHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_1()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_2()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_3()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_4()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_5()
        AddHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
        AddHandler Me.ComboBox5.SelectedIndexChanged, AddressOf ComboBox5_SelectedIndexChanged
        AddHandler Me.ComboBox4.SelectedIndexChanged, AddressOf ComboBox4_SelectedIndexChanged
        AddHandler Me.ComboBox3.SelectedIndexChanged, AddressOf ComboBox3_SelectedIndexChanged
        AddHandler Me.ComboBox2.SelectedIndexChanged, AddressOf ComboBox2_SelectedIndexChanged
        Call GENERAR_ORDEN()
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_ESTRUCTURA()
        Me.TextBox2.Focus()
        Me.TextBox2.SelectAll()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox2.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_5()
        If Me.ComboBox5.Items.Count > 0 Then
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Conexion.ABD_RRHH()
            Dim MiDataTable As New DataTable
            Dim MiSqlDataAdapter As SqlDataAdapter
            Try
                MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTRUCTURA NIVEL 5] where ID_N4 = " & Me.ComboBox5.SelectedValue & " order by ORDEN", Conexion.CxRRHH)
                MiDataTable.Clear()
                MiSqlDataAdapter.Fill(MiDataTable)
                If MiDataTable.Rows.Count <> 0 Then
                    Me.ComboBox6.Font = New Font("Microsoft Sans Serif", 8.25)
                    Me.ComboBox6.ForeColor = Color.Black
                    With Me.ComboBox6
                        .DataSource = MiDataTable
                        .DisplayMember = "DESCRIPCION"
                        .ValueMember = "ID_N5"
                    End With
                Else
                    Me.ComboBox6.DataSource = Nothing
                    Me.ComboBox6.ForeColor = Color.Red
                    Me.ComboBox6.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
                    Me.ComboBox6.Text = "-- SIN DATOS -- "
                End If
            Catch ex As Exception
                MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
        Else
            Me.ComboBox6.DataSource = Nothing
            Me.ComboBox6.ForeColor = Color.Red
            Me.ComboBox6.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
            Me.ComboBox6.Text = "-- SIN DATOS -- "
        End If
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_4()
        If Me.ComboBox4.Items.Count > 0 Then
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Conexion.ABD_RRHH()
            Dim MiDataTable As New DataTable
            Dim MiSqlDataAdapter As SqlDataAdapter
            Try
                MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTRUCTURA NIVEL 4] where ID_N3 = " & Me.ComboBox4.SelectedValue & " order by ORDEN", Conexion.CxRRHH)
                MiDataTable.Clear()
                MiSqlDataAdapter.Fill(MiDataTable)
                If MiDataTable.Rows.Count <> 0 Then
                    Me.ComboBox5.Font = New Font("Microsoft Sans Serif", 8.25)
                    Me.ComboBox5.ForeColor = Color.Black
                    With Me.ComboBox5
                        .DataSource = MiDataTable
                        .DisplayMember = "DESCRIPCION"
                        .ValueMember = "ID_N4"
                    End With
                Else
                    Me.ComboBox5.DataSource = Nothing
                    Me.ComboBox5.ForeColor = Color.Red
                    Me.ComboBox5.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
                    Me.ComboBox5.Text = "-- SIN DATOS -- "
                End If
            Catch ex As Exception
                MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
        Else
            Me.ComboBox5.DataSource = Nothing
            Me.ComboBox5.ForeColor = Color.Red
            Me.ComboBox5.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
            Me.ComboBox5.Text = "-- SIN DATOS -- "
        End If
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_3()
        If Me.ComboBox3.Items.Count > 0 Then
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Conexion.ABD_RRHH()
            Dim MiDataTable As New DataTable
            Dim MiSqlDataAdapter As SqlDataAdapter
            Try
                MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTRUCTURA NIVEL 3] where ID_N2 = " & Me.ComboBox3.SelectedValue & " order by ORDEN", Conexion.CxRRHH)
                MiDataTable.Clear()
                MiSqlDataAdapter.Fill(MiDataTable)
                If MiDataTable.Rows.Count <> 0 Then
                    Me.ComboBox4.Font = New Font("Microsoft Sans Serif", 8.25)
                    Me.ComboBox4.ForeColor = Color.Black
                    With Me.ComboBox4
                        .DataSource = MiDataTable
                        .DisplayMember = "DESCRIPCION"
                        .ValueMember = "ID_N3"
                    End With
                Else
                    Me.ComboBox4.DataSource = Nothing
                    Me.ComboBox4.ForeColor = Color.Red
                    Me.ComboBox4.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
                    Me.ComboBox4.Text = "-- SIN DATOS -- "
                End If
            Catch ex As Exception
                MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
        Else
            Me.ComboBox4.DataSource = Nothing
            Me.ComboBox4.ForeColor = Color.Red
            Me.ComboBox4.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
            Me.ComboBox4.Text = "-- SIN DATOS -- "
        End If
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_2()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTRUCTURA NIVEL 2] where ID_N1 = " & Me.ComboBox2.SelectedValue & " order by ORDEN", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count <> 0 Then
                Me.ComboBox3.Font = New Font("Microsoft Sans Serif", 8.25)
                Me.ComboBox3.ForeColor = Color.Black
                With Me.ComboBox3
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_N2"
                End With
            Else
                Me.ComboBox3.DataSource = Nothing
                Me.ComboBox3.ForeColor = Color.Red
                Me.ComboBox3.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
                Me.ComboBox3.Text = "-- SIN DATOS --"
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
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
    Private Sub ComboBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub GENERAR_ORDEN()
        CODIGO = 0
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            MiDataAdapter = New SqlDataAdapter("SELECT ORDEN FROM [CAT DE ESTRUCTURA NIVEL 6] ORDER BY ORDEN", Conexion.CxRRHH)
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
            Me.TextBox2.Text = CODIGO
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
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
            MiDataAdapter = New SqlDataAdapter("SELECT ID_N6 FROM [CAT DE ESTRUCTURA NIVEL 6] ORDER BY ID_N6", Conexion.CxRRHH)
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
        If Len(Me.ComboBox2.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Nivel 1 Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox2.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox3.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Nivel 2 Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox3.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox4.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Nivel 3 Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox4.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox5.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Nivel 4 Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox5.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox6.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Nivel 5 Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox6.Focus()
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
        Call GRABAR_REGISTRO()
        Call REENUMERAR()
        n6ID_N6 = Me.TextBox1.Text
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
        Call SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        Me.TextBox1.Text = ""
        Call GENERAR_ORDEN()
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        'Me.TextBox2.Focus()
        'Me.TextBox2.SelectAll()
        Me.ComboBox4.Focus()
        Me.ComboBox4.SelectAll()
        CERRAR = False
        'Me.Close()
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
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [dbo].[VISTA CAT DE ESTRUCTURA NIVEL 6] ORDER BY [ORD6]", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                For I = 0 To MiDataTable.Rows.Count - 1
                    CONTADOR = CONTADOR + 1
                    iINDICE = MiDataTable.Rows(I).Item(15).ToString
                    iNORDEN = CONTADOR
                    Call ACTUALIZAR_CAT_DE_ESTRUCTURA_NIVEL_6()
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
    Private Sub ACTUALIZAR_CAT_DE_ESTRUCTURA_NIVEL_6()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [CAT DE ESTRUCTURA NIVEL 6] SET ORDEN = '" & iNORDEN & "' WHERE ID_N6 = " & CInt(iINDICE) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub GRABAR_REGISTRO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [CAT DE ESTRUCTURA NIVEL 6] (ID_N6, ID_N5, ORDEN, DESCRIPCION, SIGLAS, ID_T_ES, MIXTA, MILITAR, PAME)" &
                                  "values (" & CInt(Me.TextBox1.Text) & ", '" & Me.ComboBox6.SelectedValue & "', '" & Trim(Me.TextBox2.Text) & "', '" & Trim(Me.TextBox3.Text) & "', '" & Trim(Me.TextBox4.Text) & "', " & Me.ComboBox1.SelectedValue & ", '" & Me.CheckBox1.Checked & "', '" & Me.CheckBox2.Checked & "', '" & Me.CheckBox3.Checked & "')"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        RemoveHandler Me.ComboBox3.SelectedIndexChanged, AddressOf ComboBox3_SelectedIndexChanged
        RemoveHandler Me.ComboBox4.SelectedIndexChanged, AddressOf ComboBox4_SelectedIndexChanged
        RemoveHandler Me.ComboBox5.SelectedIndexChanged, AddressOf ComboBox5_SelectedIndexChanged
        RemoveHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_2()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_3()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_4()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_5()
        AddHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
        AddHandler Me.ComboBox5.SelectedIndexChanged, AddressOf ComboBox5_SelectedIndexChanged
        AddHandler Me.ComboBox4.SelectedIndexChanged, AddressOf ComboBox4_SelectedIndexChanged
        AddHandler Me.ComboBox3.SelectedIndexChanged, AddressOf ComboBox3_SelectedIndexChanged
    End Sub
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        RemoveHandler Me.ComboBox4.SelectedIndexChanged, AddressOf ComboBox4_SelectedIndexChanged
        RemoveHandler Me.ComboBox5.SelectedIndexChanged, AddressOf ComboBox5_SelectedIndexChanged
        RemoveHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_3()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_4()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_5()
        AddHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
        AddHandler Me.ComboBox5.SelectedIndexChanged, AddressOf ComboBox5_SelectedIndexChanged
        AddHandler Me.ComboBox4.SelectedIndexChanged, AddressOf ComboBox4_SelectedIndexChanged
    End Sub
    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        RemoveHandler Me.ComboBox5.SelectedIndexChanged, AddressOf ComboBox5_SelectedIndexChanged
        RemoveHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_4()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_5()
        AddHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
        AddHandler Me.ComboBox5.SelectedIndexChanged, AddressOf ComboBox5_SelectedIndexChanged
    End Sub
    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        RemoveHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_5()
        AddHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
    End Sub
    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged

    End Sub
    Private Sub CARGAR_DATAGRIDVIEW()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter("SELECT * FROM [dbo].[VISTA CAT DE ESTRUCTURA NIVEL 6] ORDER BY [ORD6]", Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA CAT DE ESTRUCTURA NIVEL 6]")
        Frm_Cat_Estructura_N6.DGV.DataSource = MiDataSet.Tables("[VISTA CAT DE ESTRUCTURA NIVEL 6]")
        Frm_Cat_Estructura_N6.Label2.Text = "Total de Registros: " & Frm_Cat_Estructura_N6.DGV.RowCount
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW()
        Frm_Cat_Estructura_N6.DGV.Columns(0).HeaderText = "ID_N1"
        Frm_Cat_Estructura_N6.DGV.Columns(0).Width = 0
        Frm_Cat_Estructura_N6.DGV.Columns(0).Visible = False

        Frm_Cat_Estructura_N6.DGV.Columns(1).HeaderText = "ORD1"
        Frm_Cat_Estructura_N6.DGV.Columns(1).Width = 0
        Frm_Cat_Estructura_N6.DGV.Columns(1).Visible = False

        Frm_Cat_Estructura_N6.DGV.Columns(2).HeaderText = "Descripcion N1"
        Frm_Cat_Estructura_N6.DGV.Columns(2).Width = 100
        Frm_Cat_Estructura_N6.DGV.Columns(2).Visible = True
        Frm_Cat_Estructura_N6.DGV.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        'Frm_Cat_Estructura_N6.DGV.Columns(2).Frozen = True

        Frm_Cat_Estructura_N6.DGV.Columns(3).HeaderText = "ID_N2"
        Frm_Cat_Estructura_N6.DGV.Columns(3).Width = 0
        Frm_Cat_Estructura_N6.DGV.Columns(3).Visible = False

        Frm_Cat_Estructura_N6.DGV.Columns(4).HeaderText = "ORDN2"
        Frm_Cat_Estructura_N6.DGV.Columns(4).Width = 0
        Frm_Cat_Estructura_N6.DGV.Columns(4).Visible = False

        Frm_Cat_Estructura_N6.DGV.Columns(5).HeaderText = "Descripcion N2"
        Frm_Cat_Estructura_N6.DGV.Columns(5).Width = 100
        Frm_Cat_Estructura_N6.DGV.Columns(5).Visible = True
        Frm_Cat_Estructura_N6.DGV.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Cat_Estructura_N6.DGV.Columns(6).HeaderText = "ID_N3"
        Frm_Cat_Estructura_N6.DGV.Columns(6).Width = 0
        Frm_Cat_Estructura_N6.DGV.Columns(6).Visible = False

        Frm_Cat_Estructura_N6.DGV.Columns(7).HeaderText = "ORDN3"
        Frm_Cat_Estructura_N6.DGV.Columns(7).Width = 0
        Frm_Cat_Estructura_N6.DGV.Columns(7).Visible = False

        Frm_Cat_Estructura_N6.DGV.Columns(8).HeaderText = "Descripcion N3"
        Frm_Cat_Estructura_N6.DGV.Columns(8).Width = 100
        Frm_Cat_Estructura_N6.DGV.Columns(8).Visible = True
        Frm_Cat_Estructura_N6.DGV.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Cat_Estructura_N6.DGV.Columns(9).HeaderText = "ID_N4"
        Frm_Cat_Estructura_N6.DGV.Columns(9).Width = 0
        Frm_Cat_Estructura_N6.DGV.Columns(9).Visible = False

        Frm_Cat_Estructura_N6.DGV.Columns(10).HeaderText = "ORDN4"
        Frm_Cat_Estructura_N6.DGV.Columns(10).Width = 0
        Frm_Cat_Estructura_N6.DGV.Columns(10).Visible = False

        Frm_Cat_Estructura_N6.DGV.Columns(11).HeaderText = "Descripcion N4"
        Frm_Cat_Estructura_N6.DGV.Columns(11).Width = 150
        Frm_Cat_Estructura_N6.DGV.Columns(11).Visible = True
        Frm_Cat_Estructura_N6.DGV.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Cat_Estructura_N6.DGV.Columns(12).HeaderText = "ID_N5"
        Frm_Cat_Estructura_N6.DGV.Columns(12).Width = 0
        Frm_Cat_Estructura_N6.DGV.Columns(12).Visible = False

        Frm_Cat_Estructura_N6.DGV.Columns(13).HeaderText = "ORDN5"
        Frm_Cat_Estructura_N6.DGV.Columns(13).Width = 0
        Frm_Cat_Estructura_N6.DGV.Columns(13).Visible = False

        Frm_Cat_Estructura_N6.DGV.Columns(14).HeaderText = "Descripcion N5"
        Frm_Cat_Estructura_N6.DGV.Columns(14).Width = 150
        Frm_Cat_Estructura_N6.DGV.Columns(14).Visible = True
        Frm_Cat_Estructura_N6.DGV.Columns(14).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Cat_Estructura_N6.DGV.Columns(15).HeaderText = "Id"
        Frm_Cat_Estructura_N6.DGV.Columns(15).Width = 10
        Frm_Cat_Estructura_N6.DGV.Columns(15).Visible = True
        Frm_Cat_Estructura_N6.DGV.Columns(15).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Cat_Estructura_N6.DGV.Columns(15).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Frm_Cat_Estructura_N6.DGV.Columns(15).DefaultCellStyle.BackColor = Color.Black

        Frm_Cat_Estructura_N6.DGV.Columns(16).HeaderText = "Orden"
        Frm_Cat_Estructura_N6.DGV.Columns(16).Width = 40
        Frm_Cat_Estructura_N6.DGV.Columns(16).Visible = True
        Frm_Cat_Estructura_N6.DGV.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Cat_Estructura_N6.DGV.Columns(16).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Frm_Cat_Estructura_N6.DGV.Columns(16).DefaultCellStyle.BackColor = Color.Khaki

        Frm_Cat_Estructura_N6.DGV.Columns(17).HeaderText = "Descripcion N6"
        Frm_Cat_Estructura_N6.DGV.Columns(17).Width = 160
        Frm_Cat_Estructura_N6.DGV.Columns(17).Visible = True
        Frm_Cat_Estructura_N6.DGV.Columns(17).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Cat_Estructura_N6.DGV.Columns(17).DefaultCellStyle.BackColor = Color.LightGreen

        Frm_Cat_Estructura_N6.DGV.Columns(18).HeaderText = "Siglas"
        Frm_Cat_Estructura_N6.DGV.Columns(18).Width = 60
        Frm_Cat_Estructura_N6.DGV.Columns(18).Visible = True
        Frm_Cat_Estructura_N6.DGV.Columns(18).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Cat_Estructura_N6.DGV.Columns(19).HeaderText = "ID_T_ES"
        Frm_Cat_Estructura_N6.DGV.Columns(19).Width = 0
        Frm_Cat_Estructura_N6.DGV.Columns(19).Visible = False
        Frm_Cat_Estructura_N6.DGV.Columns(19).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Cat_Estructura_N6.DGV.Columns(20).HeaderText = "Tipo"
        Frm_Cat_Estructura_N6.DGV.Columns(20).Width = 80
        Frm_Cat_Estructura_N6.DGV.Columns(20).Visible = True
        Frm_Cat_Estructura_N6.DGV.Columns(20).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Cat_Estructura_N6.DGV.Columns(21).HeaderText = "Mixta"
        Frm_Cat_Estructura_N6.DGV.Columns(21).Width = 40
        Frm_Cat_Estructura_N6.DGV.Columns(21).Visible = True
        Frm_Cat_Estructura_N6.DGV.Columns(21).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Cat_Estructura_N6.DGV.Columns(22).HeaderText = "Militar"
        Frm_Cat_Estructura_N6.DGV.Columns(22).Width = 40
        Frm_Cat_Estructura_N6.DGV.Columns(22).Visible = True
        Frm_Cat_Estructura_N6.DGV.Columns(22).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Cat_Estructura_N6.DGV.Columns(23).HeaderText = "Pame"
        Frm_Cat_Estructura_N6.DGV.Columns(23).Width = 40
        Frm_Cat_Estructura_N6.DGV.Columns(23).Visible = True
        Frm_Cat_Estructura_N6.DGV.Columns(23).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub SELECCIONAR_REGISTRO_DATAGRIDVIEW()
        Dim I As Integer
        For I = 0 To Frm_Cat_Estructura_N6.DGV.RowCount - 1
            If Frm_Cat_Estructura_N6.DGV.Item(15, I).Value = n6ID_N6 Then
                Frm_Cat_Estructura_N6.DGV.Rows(I).Selected = True
                Frm_Cat_Estructura_N6.DGV.CurrentCell = Frm_Cat_Estructura_N6.DGV.Item(15, I)
                Exit Sub
            End If
        Next
    End Sub
End Class