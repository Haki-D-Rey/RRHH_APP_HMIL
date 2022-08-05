Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Personal_Ausente_Dias_Continuos
    Private Sub Frm_Personal_Ausente_Dias_Continuos_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            CONSULTA_PERSONAL_AUSENTE = False
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Personal_Ausente_Dias_Continuos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CONSULTA_PERSONAL_AUSENTE = True
        Call Clases.BUSCAR_FECHA_Y_HORA()
        Me.DateTimePicker2.Value = Mid(FECHA_SERVIDOR, 1, 10)
        Me.DateTimePicker1.Value = Me.DateTimePicker2.Value.AddDays(-4)
        RemoveHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
        RemoveHandler Me.ComboBox2.SelectedIndexChanged, AddressOf ComboBox2_SelectedIndexChanged
        RemoveHandler Me.ComboBox3.SelectedIndexChanged, AddressOf ComboBox3_SelectedIndexChanged
        RemoveHandler Me.ComboBox4.SelectedIndexChanged, AddressOf ComboBox4_SelectedIndexChanged
        RemoveHandler Me.ComboBox5.SelectedIndexChanged, AddressOf ComboBox5_SelectedIndexChanged
        RemoveHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
        RemoveHandler Me.ComboBox7.SelectedIndexChanged, AddressOf ComboBox7_SelectedIndexChanged
        RemoveHandler Me.ComboBox8.SelectedIndexChanged, AddressOf ComboBox8_SelectedIndexChanged
        RemoveHandler CheckBox1.CheckedChanged, AddressOf CheckBox1_CheckedChanged
        RemoveHandler CheckBox2.CheckedChanged, AddressOf CheckBox2_CheckedChanged
        RemoveHandler CheckBox3.CheckedChanged, AddressOf CheckBox3_CheckedChanged
        RemoveHandler CheckBox4.CheckedChanged, AddressOf CheckBox4_CheckedChanged
        RemoveHandler CheckBox5.CheckedChanged, AddressOf CheckBox5_CheckedChanged
        RemoveHandler CheckBox6.CheckedChanged, AddressOf CheckBox6_CheckedChanged
        RemoveHandler CheckBox7.CheckedChanged, AddressOf CheckBox7_CheckedChanged
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_1()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_2()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_3()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_4()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_5()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_6()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_7()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_8()
        AddHandler Me.CheckBox7.CheckedChanged, AddressOf CheckBox7_CheckedChanged
        AddHandler Me.CheckBox6.CheckedChanged, AddressOf CheckBox6_CheckedChanged
        AddHandler Me.CheckBox5.CheckedChanged, AddressOf CheckBox5_CheckedChanged
        AddHandler Me.CheckBox4.CheckedChanged, AddressOf CheckBox4_CheckedChanged
        AddHandler Me.CheckBox3.CheckedChanged, AddressOf CheckBox3_CheckedChanged
        AddHandler Me.CheckBox2.CheckedChanged, AddressOf CheckBox2_CheckedChanged
        AddHandler Me.CheckBox1.CheckedChanged, AddressOf CheckBox1_CheckedChanged
        AddHandler Me.ComboBox8.SelectedIndexChanged, AddressOf ComboBox8_SelectedIndexChanged
        AddHandler Me.ComboBox7.SelectedIndexChanged, AddressOf ComboBox7_SelectedIndexChanged
        AddHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
        AddHandler Me.ComboBox5.SelectedIndexChanged, AddressOf ComboBox5_SelectedIndexChanged
        AddHandler Me.ComboBox4.SelectedIndexChanged, AddressOf ComboBox4_SelectedIndexChanged
        AddHandler Me.ComboBox3.SelectedIndexChanged, AddressOf ComboBox3_SelectedIndexChanged
        AddHandler Me.ComboBox2.SelectedIndexChanged, AddressOf ComboBox2_SelectedIndexChanged
        AddHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        RemoveHandler Me.ComboBox2.SelectedIndexChanged, AddressOf ComboBox2_SelectedIndexChanged
        RemoveHandler Me.ComboBox3.SelectedIndexChanged, AddressOf ComboBox3_SelectedIndexChanged
        RemoveHandler Me.ComboBox4.SelectedIndexChanged, AddressOf ComboBox4_SelectedIndexChanged
        RemoveHandler Me.ComboBox5.SelectedIndexChanged, AddressOf ComboBox5_SelectedIndexChanged
        RemoveHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
        RemoveHandler Me.ComboBox7.SelectedIndexChanged, AddressOf ComboBox7_SelectedIndexChanged
        RemoveHandler Me.ComboBox8.SelectedIndexChanged, AddressOf ComboBox8_SelectedIndexChanged
        RemoveHandler CheckBox1.CheckedChanged, AddressOf CheckBox1_CheckedChanged
        RemoveHandler CheckBox2.CheckedChanged, AddressOf CheckBox2_CheckedChanged
        RemoveHandler CheckBox3.CheckedChanged, AddressOf CheckBox3_CheckedChanged
        RemoveHandler CheckBox4.CheckedChanged, AddressOf CheckBox4_CheckedChanged
        RemoveHandler CheckBox5.CheckedChanged, AddressOf CheckBox5_CheckedChanged
        RemoveHandler CheckBox6.CheckedChanged, AddressOf CheckBox6_CheckedChanged
        RemoveHandler CheckBox7.CheckedChanged, AddressOf CheckBox7_CheckedChanged
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_2()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_3()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_4()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_5()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_6()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_7()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_8()
        AddHandler Me.CheckBox7.CheckedChanged, AddressOf CheckBox7_CheckedChanged
        AddHandler Me.CheckBox6.CheckedChanged, AddressOf CheckBox6_CheckedChanged
        AddHandler Me.CheckBox5.CheckedChanged, AddressOf CheckBox5_CheckedChanged
        AddHandler Me.CheckBox4.CheckedChanged, AddressOf CheckBox4_CheckedChanged
        AddHandler Me.CheckBox3.CheckedChanged, AddressOf CheckBox3_CheckedChanged
        AddHandler Me.CheckBox2.CheckedChanged, AddressOf CheckBox2_CheckedChanged
        AddHandler Me.CheckBox1.CheckedChanged, AddressOf CheckBox1_CheckedChanged
        AddHandler Me.ComboBox8.SelectedIndexChanged, AddressOf ComboBox8_SelectedIndexChanged
        AddHandler Me.ComboBox7.SelectedIndexChanged, AddressOf ComboBox7_SelectedIndexChanged
        AddHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
        AddHandler Me.ComboBox5.SelectedIndexChanged, AddressOf ComboBox5_SelectedIndexChanged
        AddHandler Me.ComboBox4.SelectedIndexChanged, AddressOf ComboBox4_SelectedIndexChanged
        AddHandler Me.ComboBox3.SelectedIndexChanged, AddressOf ComboBox3_SelectedIndexChanged
        AddHandler Me.ComboBox2.SelectedIndexChanged, AddressOf ComboBox2_SelectedIndexChanged
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        RemoveHandler Me.ComboBox3.SelectedIndexChanged, AddressOf ComboBox3_SelectedIndexChanged
        RemoveHandler Me.ComboBox4.SelectedIndexChanged, AddressOf ComboBox4_SelectedIndexChanged
        RemoveHandler Me.ComboBox5.SelectedIndexChanged, AddressOf ComboBox5_SelectedIndexChanged
        RemoveHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
        RemoveHandler Me.ComboBox7.SelectedIndexChanged, AddressOf ComboBox7_SelectedIndexChanged
        RemoveHandler Me.ComboBox8.SelectedIndexChanged, AddressOf ComboBox8_SelectedIndexChanged
        RemoveHandler CheckBox1.CheckedChanged, AddressOf CheckBox1_CheckedChanged
        RemoveHandler CheckBox2.CheckedChanged, AddressOf CheckBox2_CheckedChanged
        RemoveHandler CheckBox3.CheckedChanged, AddressOf CheckBox3_CheckedChanged
        RemoveHandler CheckBox4.CheckedChanged, AddressOf CheckBox4_CheckedChanged
        RemoveHandler CheckBox5.CheckedChanged, AddressOf CheckBox5_CheckedChanged
        RemoveHandler CheckBox6.CheckedChanged, AddressOf CheckBox6_CheckedChanged
        RemoveHandler CheckBox7.CheckedChanged, AddressOf CheckBox7_CheckedChanged
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_3()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_4()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_5()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_6()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_7()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_8()
        AddHandler Me.CheckBox7.CheckedChanged, AddressOf CheckBox7_CheckedChanged
        AddHandler Me.CheckBox6.CheckedChanged, AddressOf CheckBox6_CheckedChanged
        AddHandler Me.CheckBox5.CheckedChanged, AddressOf CheckBox5_CheckedChanged
        AddHandler Me.CheckBox4.CheckedChanged, AddressOf CheckBox4_CheckedChanged
        AddHandler Me.CheckBox3.CheckedChanged, AddressOf CheckBox3_CheckedChanged
        AddHandler Me.CheckBox2.CheckedChanged, AddressOf CheckBox2_CheckedChanged
        AddHandler Me.CheckBox1.CheckedChanged, AddressOf CheckBox1_CheckedChanged
        AddHandler Me.ComboBox8.SelectedIndexChanged, AddressOf ComboBox8_SelectedIndexChanged
        AddHandler Me.ComboBox7.SelectedIndexChanged, AddressOf ComboBox7_SelectedIndexChanged
        AddHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
        AddHandler Me.ComboBox5.SelectedIndexChanged, AddressOf ComboBox5_SelectedIndexChanged
        AddHandler Me.ComboBox4.SelectedIndexChanged, AddressOf ComboBox4_SelectedIndexChanged
        AddHandler Me.ComboBox3.SelectedIndexChanged, AddressOf ComboBox3_SelectedIndexChanged
    End Sub
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        RemoveHandler Me.ComboBox4.SelectedIndexChanged, AddressOf ComboBox4_SelectedIndexChanged
        RemoveHandler Me.ComboBox5.SelectedIndexChanged, AddressOf ComboBox5_SelectedIndexChanged
        RemoveHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
        RemoveHandler Me.ComboBox7.SelectedIndexChanged, AddressOf ComboBox7_SelectedIndexChanged
        RemoveHandler Me.ComboBox8.SelectedIndexChanged, AddressOf ComboBox8_SelectedIndexChanged
        RemoveHandler CheckBox1.CheckedChanged, AddressOf CheckBox1_CheckedChanged
        RemoveHandler CheckBox2.CheckedChanged, AddressOf CheckBox2_CheckedChanged
        RemoveHandler CheckBox3.CheckedChanged, AddressOf CheckBox3_CheckedChanged
        RemoveHandler CheckBox4.CheckedChanged, AddressOf CheckBox4_CheckedChanged
        RemoveHandler CheckBox5.CheckedChanged, AddressOf CheckBox5_CheckedChanged
        RemoveHandler CheckBox6.CheckedChanged, AddressOf CheckBox6_CheckedChanged
        RemoveHandler CheckBox7.CheckedChanged, AddressOf CheckBox7_CheckedChanged
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_4()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_5()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_6()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_7()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_8()
        AddHandler Me.CheckBox7.CheckedChanged, AddressOf CheckBox7_CheckedChanged
        AddHandler Me.CheckBox6.CheckedChanged, AddressOf CheckBox6_CheckedChanged
        AddHandler Me.CheckBox5.CheckedChanged, AddressOf CheckBox5_CheckedChanged
        AddHandler Me.CheckBox4.CheckedChanged, AddressOf CheckBox4_CheckedChanged
        AddHandler Me.CheckBox3.CheckedChanged, AddressOf CheckBox3_CheckedChanged
        AddHandler Me.CheckBox2.CheckedChanged, AddressOf CheckBox2_CheckedChanged
        AddHandler Me.CheckBox1.CheckedChanged, AddressOf CheckBox1_CheckedChanged
        AddHandler Me.ComboBox8.SelectedIndexChanged, AddressOf ComboBox8_SelectedIndexChanged
        AddHandler Me.ComboBox7.SelectedIndexChanged, AddressOf ComboBox7_SelectedIndexChanged
        AddHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
        AddHandler Me.ComboBox5.SelectedIndexChanged, AddressOf ComboBox5_SelectedIndexChanged
        AddHandler Me.ComboBox4.SelectedIndexChanged, AddressOf ComboBox4_SelectedIndexChanged
    End Sub
    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        RemoveHandler Me.ComboBox5.SelectedIndexChanged, AddressOf ComboBox5_SelectedIndexChanged
        RemoveHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
        RemoveHandler Me.ComboBox7.SelectedIndexChanged, AddressOf ComboBox7_SelectedIndexChanged
        RemoveHandler Me.ComboBox8.SelectedIndexChanged, AddressOf ComboBox8_SelectedIndexChanged
        RemoveHandler CheckBox1.CheckedChanged, AddressOf CheckBox1_CheckedChanged
        RemoveHandler CheckBox2.CheckedChanged, AddressOf CheckBox2_CheckedChanged
        RemoveHandler CheckBox3.CheckedChanged, AddressOf CheckBox3_CheckedChanged
        RemoveHandler CheckBox4.CheckedChanged, AddressOf CheckBox4_CheckedChanged
        RemoveHandler CheckBox5.CheckedChanged, AddressOf CheckBox5_CheckedChanged
        RemoveHandler CheckBox6.CheckedChanged, AddressOf CheckBox6_CheckedChanged
        RemoveHandler CheckBox7.CheckedChanged, AddressOf CheckBox7_CheckedChanged
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_5()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_6()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_7()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_8()
        AddHandler Me.CheckBox7.CheckedChanged, AddressOf CheckBox7_CheckedChanged
        AddHandler Me.CheckBox6.CheckedChanged, AddressOf CheckBox6_CheckedChanged
        AddHandler Me.CheckBox5.CheckedChanged, AddressOf CheckBox5_CheckedChanged
        AddHandler Me.CheckBox4.CheckedChanged, AddressOf CheckBox4_CheckedChanged
        AddHandler Me.CheckBox3.CheckedChanged, AddressOf CheckBox3_CheckedChanged
        AddHandler Me.CheckBox2.CheckedChanged, AddressOf CheckBox2_CheckedChanged
        AddHandler Me.CheckBox1.CheckedChanged, AddressOf CheckBox1_CheckedChanged
        AddHandler Me.ComboBox8.SelectedIndexChanged, AddressOf ComboBox8_SelectedIndexChanged
        AddHandler Me.ComboBox7.SelectedIndexChanged, AddressOf ComboBox7_SelectedIndexChanged
        AddHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
        AddHandler Me.ComboBox5.SelectedIndexChanged, AddressOf ComboBox5_SelectedIndexChanged
    End Sub
    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        RemoveHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
        RemoveHandler Me.ComboBox7.SelectedIndexChanged, AddressOf ComboBox7_SelectedIndexChanged
        RemoveHandler Me.ComboBox8.SelectedIndexChanged, AddressOf ComboBox8_SelectedIndexChanged
        RemoveHandler CheckBox1.CheckedChanged, AddressOf CheckBox1_CheckedChanged
        RemoveHandler CheckBox2.CheckedChanged, AddressOf CheckBox2_CheckedChanged
        RemoveHandler CheckBox3.CheckedChanged, AddressOf CheckBox3_CheckedChanged
        RemoveHandler CheckBox4.CheckedChanged, AddressOf CheckBox4_CheckedChanged
        RemoveHandler CheckBox5.CheckedChanged, AddressOf CheckBox5_CheckedChanged
        RemoveHandler CheckBox6.CheckedChanged, AddressOf CheckBox6_CheckedChanged
        RemoveHandler CheckBox7.CheckedChanged, AddressOf CheckBox7_CheckedChanged
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_6()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_7()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_8()
        AddHandler Me.CheckBox7.CheckedChanged, AddressOf CheckBox7_CheckedChanged
        AddHandler Me.CheckBox6.CheckedChanged, AddressOf CheckBox6_CheckedChanged
        AddHandler Me.CheckBox5.CheckedChanged, AddressOf CheckBox5_CheckedChanged
        AddHandler Me.CheckBox4.CheckedChanged, AddressOf CheckBox4_CheckedChanged
        AddHandler Me.CheckBox3.CheckedChanged, AddressOf CheckBox3_CheckedChanged
        AddHandler Me.CheckBox2.CheckedChanged, AddressOf CheckBox2_CheckedChanged
        AddHandler Me.CheckBox1.CheckedChanged, AddressOf CheckBox1_CheckedChanged
        AddHandler Me.ComboBox8.SelectedIndexChanged, AddressOf ComboBox8_SelectedIndexChanged
        AddHandler Me.ComboBox7.SelectedIndexChanged, AddressOf ComboBox7_SelectedIndexChanged
        AddHandler Me.ComboBox6.SelectedIndexChanged, AddressOf ComboBox6_SelectedIndexChanged
    End Sub
    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        RemoveHandler Me.ComboBox7.SelectedIndexChanged, AddressOf ComboBox7_SelectedIndexChanged
        RemoveHandler Me.ComboBox8.SelectedIndexChanged, AddressOf ComboBox8_SelectedIndexChanged
        RemoveHandler CheckBox1.CheckedChanged, AddressOf CheckBox1_CheckedChanged
        RemoveHandler CheckBox2.CheckedChanged, AddressOf CheckBox2_CheckedChanged
        RemoveHandler CheckBox3.CheckedChanged, AddressOf CheckBox3_CheckedChanged
        RemoveHandler CheckBox4.CheckedChanged, AddressOf CheckBox4_CheckedChanged
        RemoveHandler CheckBox5.CheckedChanged, AddressOf CheckBox5_CheckedChanged
        RemoveHandler CheckBox6.CheckedChanged, AddressOf CheckBox6_CheckedChanged
        RemoveHandler CheckBox7.CheckedChanged, AddressOf CheckBox7_CheckedChanged
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_7()
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_8()
        AddHandler Me.CheckBox7.CheckedChanged, AddressOf CheckBox7_CheckedChanged
        AddHandler Me.CheckBox6.CheckedChanged, AddressOf CheckBox6_CheckedChanged
        AddHandler Me.CheckBox5.CheckedChanged, AddressOf CheckBox5_CheckedChanged
        AddHandler Me.CheckBox4.CheckedChanged, AddressOf CheckBox4_CheckedChanged
        AddHandler Me.CheckBox3.CheckedChanged, AddressOf CheckBox3_CheckedChanged
        AddHandler Me.CheckBox2.CheckedChanged, AddressOf CheckBox2_CheckedChanged
        AddHandler Me.CheckBox1.CheckedChanged, AddressOf CheckBox1_CheckedChanged
        AddHandler Me.ComboBox8.SelectedIndexChanged, AddressOf ComboBox8_SelectedIndexChanged
        AddHandler Me.ComboBox7.SelectedIndexChanged, AddressOf ComboBox7_SelectedIndexChanged
    End Sub
    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        RemoveHandler Me.ComboBox8.SelectedIndexChanged, AddressOf ComboBox8_SelectedIndexChanged
        RemoveHandler CheckBox1.CheckedChanged, AddressOf CheckBox1_CheckedChanged
        RemoveHandler CheckBox2.CheckedChanged, AddressOf CheckBox2_CheckedChanged
        RemoveHandler CheckBox3.CheckedChanged, AddressOf CheckBox3_CheckedChanged
        RemoveHandler CheckBox4.CheckedChanged, AddressOf CheckBox4_CheckedChanged
        RemoveHandler CheckBox5.CheckedChanged, AddressOf CheckBox5_CheckedChanged
        RemoveHandler CheckBox6.CheckedChanged, AddressOf CheckBox6_CheckedChanged
        RemoveHandler CheckBox7.CheckedChanged, AddressOf CheckBox7_CheckedChanged
        Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_8()
        AddHandler Me.CheckBox7.CheckedChanged, AddressOf CheckBox7_CheckedChanged
        AddHandler Me.CheckBox6.CheckedChanged, AddressOf CheckBox6_CheckedChanged
        AddHandler Me.CheckBox5.CheckedChanged, AddressOf CheckBox5_CheckedChanged
        AddHandler Me.CheckBox4.CheckedChanged, AddressOf CheckBox4_CheckedChanged
        AddHandler Me.CheckBox3.CheckedChanged, AddressOf CheckBox3_CheckedChanged
        AddHandler Me.CheckBox2.CheckedChanged, AddressOf CheckBox2_CheckedChanged
        AddHandler Me.CheckBox1.CheckedChanged, AddressOf CheckBox1_CheckedChanged
        AddHandler Me.ComboBox8.SelectedIndexChanged, AddressOf ComboBox8_SelectedIndexChanged
    End Sub
    Private Sub ComboBox8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox8.SelectedIndexChanged
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If Me.CheckBox1.Checked = True Then
            Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_2()
            NIVEL2sql = Me.ComboBox2.SelectedValue
        Else
            Me.ComboBox2.DataSource = Nothing
            Me.ComboBox2.ForeColor = Color.Red
            Me.ComboBox2.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
            Me.ComboBox2.Text = "-- SIN DATOS -- "
            NIVEL2sql = Nothing
            Me.CheckBox2.Checked = False
            Me.CheckBox3.Checked = False
            Me.CheckBox4.Checked = False
        End If
    End Sub
    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If Me.CheckBox2.Checked = True Then
            Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_3()
            NIVEL3sql = Me.ComboBox3.SelectedValue
        Else
            Me.ComboBox3.DataSource = Nothing
            Me.ComboBox3.ForeColor = Color.Red
            Me.ComboBox3.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
            Me.ComboBox3.Text = "-- SIN DATOS -- "
            NIVEL3sql = Nothing
            Me.CheckBox3.Checked = False
            Me.CheckBox4.Checked = False
        End If
    End Sub
    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If Me.CheckBox3.Checked = True Then
            Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_4()
            NIVEL4sql = Me.ComboBox4.SelectedValue
        Else
            Me.ComboBox4.DataSource = Nothing
            Me.ComboBox4.ForeColor = Color.Red
            Me.ComboBox4.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
            Me.ComboBox4.Text = "-- SIN DATOS -- "
            NIVEL4sql = Nothing
            Me.CheckBox4.Checked = False
        End If
    End Sub
    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If Me.CheckBox4.Checked = True Then
            Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_5()
            NIVEL5sql = Me.ComboBox5.SelectedValue
        Else
            Me.ComboBox5.DataSource = Nothing
            Me.ComboBox5.ForeColor = Color.Red
            Me.ComboBox5.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
            Me.ComboBox5.Text = "-- SIN DATOS -- "
            NIVEL5sql = Nothing
        End If
    End Sub
    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If Me.CheckBox5.Checked = True Then
            Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_6()
            NIVEL6sql = Me.ComboBox6.SelectedValue
        Else
            Me.ComboBox6.DataSource = Nothing
            Me.ComboBox6.ForeColor = Color.Red
            Me.ComboBox6.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
            Me.ComboBox6.Text = "-- SIN DATOS -- "
            NIVEL6sql = Nothing
        End If
    End Sub
    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        If Me.CheckBox6.Checked = True Then
            Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_7()
            NIVEL7sql = Me.ComboBox7.SelectedValue
        Else
            Me.ComboBox7.DataSource = Nothing
            Me.ComboBox7.ForeColor = Color.Red
            Me.ComboBox7.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
            Me.ComboBox7.Text = "-- SIN DATOS -- "
            NIVEL7sql = Nothing
        End If
    End Sub
    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        If Me.CheckBox7.Checked = True Then
            Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_8()
            NIVEL8sql = Me.ComboBox8.SelectedValue
        Else
            Me.ComboBox8.DataSource = Nothing
            Me.ComboBox8.ForeColor = Color.Red
            Me.ComboBox8.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
            Me.ComboBox8.Text = "-- SIN DATOS -- "
            NIVEL8sql = Nothing
        End If
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_8()
        If Me.ComboBox7.Items.Count > 0 Then
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Conexion.ABD_RRHH()
            Dim MiDataTable As New DataTable
            Dim MiSqlDataAdapter As SqlDataAdapter
            Try
                MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTRUCTURA NIVEL 8] where ID_N7 = " & Me.ComboBox7.SelectedValue & " order by ORDEN", Conexion.CxRRHH)
                MiDataTable.Clear()
                MiSqlDataAdapter.Fill(MiDataTable)
                If MiDataTable.Rows.Count <> 0 Then
                    Me.ComboBox8.Font = New Font("Microsoft Sans Serif", 8.25)
                    Me.ComboBox8.ForeColor = Color.Black
                    Me.CheckBox7.Checked = True : Me.CheckBox7.Enabled = True
                    With Me.ComboBox8
                        .DataSource = MiDataTable
                        .DisplayMember = "DESCRIPCION"
                        .ValueMember = "ID_N8"
                    End With
                Else
                    Me.ComboBox8.DataSource = Nothing
                    Me.ComboBox8.ForeColor = Color.Red
                    Me.ComboBox8.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
                    Me.ComboBox8.Text = "-- SIN DATOS -- "
                    Me.CheckBox7.Checked = False : Me.CheckBox7.Enabled = False
                End If
            Catch ex As Exception
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
        Else
            Me.ComboBox8.DataSource = Nothing
            Me.ComboBox8.ForeColor = Color.Red
            Me.ComboBox8.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
            Me.ComboBox8.Text = "-- SIN DATOS -- "
            Me.CheckBox7.Checked = False : Me.CheckBox7.Enabled = False
        End If
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_7()
        If Me.ComboBox6.Items.Count > 0 Then
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Conexion.ABD_RRHH()
            Dim MiDataTable As New DataTable
            Dim MiSqlDataAdapter As SqlDataAdapter
            Try
                MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTRUCTURA NIVEL 7] where ID_N6 = " & Me.ComboBox6.SelectedValue & " order by ORDEN", Conexion.CxRRHH)
                MiDataTable.Clear()
                MiSqlDataAdapter.Fill(MiDataTable)
                If MiDataTable.Rows.Count <> 0 Then
                    Me.ComboBox7.Font = New Font("Microsoft Sans Serif", 8.25)
                    Me.ComboBox7.ForeColor = Color.Black
                    Me.CheckBox6.Checked = True : Me.CheckBox6.Enabled = True
                    With Me.ComboBox7
                        .DataSource = MiDataTable
                        .DisplayMember = "DESCRIPCION"
                        .ValueMember = "ID_N7"
                    End With
                Else
                    Me.ComboBox7.DataSource = Nothing
                    Me.ComboBox7.ForeColor = Color.Red
                    Me.ComboBox7.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
                    Me.ComboBox7.Text = "-- SIN DATOS -- "
                    Me.CheckBox6.Checked = False : Me.CheckBox6.Enabled = False
                End If
            Catch ex As Exception
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
        Else
            Me.ComboBox7.DataSource = Nothing
            Me.ComboBox7.ForeColor = Color.Red
            Me.ComboBox7.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
            Me.ComboBox7.Text = "-- SIN DATOS -- "
            Me.CheckBox6.Checked = False : Me.CheckBox6.Enabled = False
        End If
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_6()
        If Me.ComboBox5.Items.Count > 0 Then
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Conexion.ABD_RRHH()
            Dim MiDataTable As New DataTable
            Dim MiSqlDataAdapter As SqlDataAdapter
            Try
                MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTRUCTURA NIVEL 6] where ID_N5 = " & Me.ComboBox5.SelectedValue & " order by ORDEN", Conexion.CxRRHH)
                MiDataTable.Clear()
                MiSqlDataAdapter.Fill(MiDataTable)
                If MiDataTable.Rows.Count <> 0 Then
                    Me.ComboBox6.Font = New Font("Microsoft Sans Serif", 8.25)
                    Me.ComboBox6.ForeColor = Color.Black
                    Me.CheckBox5.Checked = True : Me.CheckBox5.Enabled = True
                    With Me.ComboBox6
                        .DataSource = MiDataTable
                        .DisplayMember = "DESCRIPCION"
                        .ValueMember = "ID_N6"
                    End With
                Else
                    Me.ComboBox6.DataSource = Nothing
                    Me.ComboBox6.ForeColor = Color.Red
                    Me.ComboBox6.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
                    Me.ComboBox6.Text = "-- SIN DATOS -- "
                    Me.CheckBox5.Checked = False : Me.CheckBox5.Enabled = False
                End If
            Catch ex As Exception
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
            Me.CheckBox5.Checked = False : Me.CheckBox5.Enabled = False
        End If
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_5()
        If Me.ComboBox4.Items.Count > 0 Then
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Conexion.ABD_RRHH()
            Dim MiDataTable As New DataTable
            Dim MiSqlDataAdapter As SqlDataAdapter
            Try
                MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTRUCTURA NIVEL 5] where ID_N4 = " & Me.ComboBox4.SelectedValue & " order by ORDEN", Conexion.CxRRHH)
                MiDataTable.Clear()
                MiSqlDataAdapter.Fill(MiDataTable)
                If MiDataTable.Rows.Count <> 0 Then
                    Me.ComboBox5.Font = New Font("Microsoft Sans Serif", 8.25)
                    Me.ComboBox5.ForeColor = Color.Black
                    Me.CheckBox4.Checked = True : Me.CheckBox4.Enabled = True
                    With Me.ComboBox5
                        .DataSource = MiDataTable
                        .DisplayMember = "DESCRIPCION"
                        .ValueMember = "ID_N5"
                    End With
                Else
                    Me.ComboBox5.DataSource = Nothing
                    Me.ComboBox5.ForeColor = Color.Red
                    Me.ComboBox5.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
                    Me.ComboBox5.Text = "-- SIN DATOS -- "
                    Me.CheckBox4.Checked = False : Me.CheckBox4.Enabled = False
                End If
            Catch ex As Exception
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
            Me.CheckBox4.Checked = False : Me.CheckBox4.Enabled = False
        End If
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_4()
        If Me.ComboBox3.Items.Count > 0 Then
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Conexion.ABD_RRHH()
            Dim MiDataTable As New DataTable
            Dim MiSqlDataAdapter As SqlDataAdapter
            Try
                MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTRUCTURA NIVEL 4] where ID_N3 = " & Me.ComboBox3.SelectedValue & " order by ORDEN", Conexion.CxRRHH)
                MiDataTable.Clear()
                MiSqlDataAdapter.Fill(MiDataTable)
                If MiDataTable.Rows.Count <> 0 Then
                    Me.ComboBox4.Font = New Font("Microsoft Sans Serif", 8.25)
                    Me.ComboBox4.ForeColor = Color.Black
                    Me.CheckBox3.Checked = True : Me.CheckBox3.Enabled = True
                    With Me.ComboBox4
                        .DataSource = MiDataTable
                        .DisplayMember = "DESCRIPCION"
                        .ValueMember = "ID_N4"
                    End With
                Else
                    Me.ComboBox4.DataSource = Nothing
                    Me.ComboBox4.ForeColor = Color.Red
                    Me.ComboBox4.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
                    Me.ComboBox4.Text = "-- SIN DATOS -- "
                    Me.CheckBox3.Checked = False : Me.CheckBox3.Enabled = False
                End If
            Catch ex As Exception
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
            Me.CheckBox3.Checked = False : Me.CheckBox3.Enabled = False
        End If
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_3()
        If Me.ComboBox2.Items.Count > 0 Then
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Conexion.ABD_RRHH()
            Dim MiDataTable As New DataTable
            Dim MiSqlDataAdapter As SqlDataAdapter
            Try
                MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTRUCTURA NIVEL 3] where ID_N2 = " & Me.ComboBox2.SelectedValue & " order by ORDEN", Conexion.CxRRHH)
                MiDataTable.Clear()
                MiSqlDataAdapter.Fill(MiDataTable)
                If MiDataTable.Rows.Count <> 0 Then
                    Me.ComboBox3.Font = New Font("Microsoft Sans Serif", 8.25)
                    Me.ComboBox3.ForeColor = Color.Black
                    Me.CheckBox2.Checked = True : Me.CheckBox2.Enabled = True
                    With Me.ComboBox3
                        .DataSource = MiDataTable
                        .DisplayMember = "DESCRIPCION"
                        .ValueMember = "ID_N3"
                    End With
                Else
                    Me.ComboBox3.DataSource = Nothing
                    Me.ComboBox3.ForeColor = Color.Red
                    Me.ComboBox3.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
                    Me.ComboBox3.Text = "-- SIN DATOS -- "
                    Me.CheckBox2.Checked = False : Me.CheckBox2.Enabled = False
                End If
            Catch ex As Exception
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
        Else
            Me.ComboBox3.DataSource = Nothing
            Me.ComboBox3.ForeColor = Color.Red
            Me.ComboBox3.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
            Me.ComboBox3.Text = "-- SIN DATOS -- "
            Me.CheckBox2.Checked = False : Me.CheckBox2.Enabled = False
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
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTRUCTURA NIVEL 2] where ID_N1 = " & Me.ComboBox1.SelectedValue & " order by ORDEN", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count <> 0 Then
                Me.ComboBox2.Font = New Font("Microsoft Sans Serif", 8.25)
                Me.ComboBox2.ForeColor = Color.Black
                Me.CheckBox1.Checked = True : Me.CheckBox1.Enabled = True
                With Me.ComboBox2
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_N2"
                End With
            Else
                Me.ComboBox2.DataSource = Nothing
                Me.ComboBox2.ForeColor = Color.Red
                Me.ComboBox2.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
                Me.ComboBox2.Text = "-- SIN DATOS --"
                Me.CheckBox1.Checked = False : Me.CheckBox1.Enabled = False
            End If
        Catch ex As Exception
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
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_N1"
            End With

        Catch ex As Exception
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        CONSULTA_PERSONAL_AUSENTE = False
        Me.Close()
    End Sub
    Private Sub Frm_Personal_Ausente_Dias_Continuos_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        CONSULTA_PERSONAL_AUSENTE = False
    End Sub
    Private Sub Frm_Personal_Ausente_Dias_Continuos_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        CONSULTA_PERSONAL_AUSENTE = False
    End Sub
    Private Sub Frm_Personal_Ausente_Dias_Continuos_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        CONSULTA_PERSONAL_AUSENTE = False
    End Sub
    Private Sub Frm_Personal_Ausente_Dias_Continuos_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        CONSULTA_PERSONAL_AUSENTE = False
    End Sub
    Dim EXISTE_DIA_EN_MARCAS As Boolean
    Dim CONTADOR_DE_AUSENCIAS As Integer
    Dim vDIA As Boolean
    Dim v1, v2, v3, v4, v5, v6, v7, v8, v9, v10, v11, v12, v13 As String
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Call VERIFICAR_ACCESOS("308") : If HAY_ACCESO = False Then : Exit Sub : End If
        Try
            TITULO_EXCEL = "PERSONAL AUSENTE EN 4 (CUATRO) DIAS CONTINUOS"
            EXPORTAR_DATOS_A_EXCEL(DGV, "DEL " & Mid(Me.DateTimePicker1.Value, 1, 10) & " AL " & Mid(Me.DateTimePicker2.Value, 1, 10))
        Catch ex As Exception
            MessageBox.Show("No se puede generar el documento Excel.", "Mensaje del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Me.DateTimePicker2.Value = Me.DateTimePicker1.Value.AddDays(4)
    End Sub
    Dim v14 As Integer
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call VERIFICAR_ACCESOS("307") : If HAY_ACCESO = False Then : Exit Sub : End If
        Dim Dias As Integer = DateAndTime.DateDiff(DateInterval.Day, DateTimePicker1.Value, DateTimePicker2.Value)
        If Dias <> 4 Then
            MsgBox("La Cantidad de dias entre el Rango de Fechas a Evaluar no puede ser mayor a 4 (Cuatro)", vbInformation, "Mensaje del Sistema")
            Me.DateTimePicker1.Focus()
            Exit Sub
        End If
        If Me.RadioButton1.Checked = False And Me.RadioButton2.Checked = False Then
            MsgBox("Debe seleccionar el Tipo de Estructura que se incluirá en la Consulta", vbInformation, "Mensaje del Sistema")
            Me.RadioButton2.Focus()
            Exit Sub
        End If
        If Me.ComboBox1.SelectedValue = 0 Then
            MsgBox("Debe seleccionar correctamente el Nivel Organizacional 1", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If Me.CheckBox1.Checked = True Then
            If Me.ComboBox2.SelectedValue = 0 Then
                MsgBox("Debe seleccionar correctamente el Nivel Organizacional 2", vbInformation, "Mensaje del Sistema")
                Me.ComboBox2.Focus()
                Exit Sub
            End If
        End If
        If Me.CheckBox2.Checked = True Then
            If Me.ComboBox3.SelectedValue = 0 Then
                MsgBox("Debe seleccionar correctamente el Nivel Organizacional 3", vbInformation, "Mensaje del Sistema")
                Me.ComboBox3.Focus()
                Exit Sub
            End If
        End If
        If Me.CheckBox3.Checked = True Then
            If Me.ComboBox4.SelectedValue = 0 Then
                MsgBox("Debe seleccionar correctamente el Nivel Organizacional 4", vbInformation, "Mensaje del Sistema")
                Me.ComboBox4.Focus()
                Exit Sub
            End If
        End If
        If Me.CheckBox4.Checked = True Then
            If Me.ComboBox5.SelectedValue = 0 Then
                MsgBox("Debe seleccionar correctamente el Nivel Organizacional 5", vbInformation, "Mensaje del Sistema")
                Me.ComboBox5.Focus()
                Exit Sub
            End If
        End If
        If Me.CheckBox5.Checked = True Then
            If Me.ComboBox6.SelectedValue = 0 Then
                MsgBox("Debe seleccionar correctamente el Nivel Organizacional 6", vbInformation, "Mensaje del Sistema")
                Me.ComboBox6.Focus()
                Exit Sub
            End If
        End If
        If Me.CheckBox6.Checked = True Then
            If Me.ComboBox7.SelectedValue = 0 Then
                MsgBox("Debe seleccionar correctamente el Nivel Organizacional 7", vbInformation, "Mensaje del Sistema")
                Me.ComboBox7.Focus()
                Exit Sub
            End If
        End If
        If Me.CheckBox7.Checked = True Then
            If Me.ComboBox8.SelectedValue = 0 Then
                MsgBox("Debe seleccionar correctamente el Nivel Organizacional 8", vbInformation, "Mensaje del Sistema")
                Me.ComboBox8.Focus()
                Exit Sub
            End If
        End If
        Dim M As String
        M = "Se dipone a iniciar el proceso de busqueda de Personal con 4 dias continuos con Omisiones de Marcas, ¿Desea Continuar?"
        MENSAJE = MsgBox(M, vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            Call ELIMINAR_TMP_CONSULTA_AUSENTES_DIAS_CONTINUOS()
            'Armar Niveles Organizacionales
            Call ARMAR_CADENA_NIVELES()
            'Armar Tipo de Estructura
            If Me.RadioButton1.Checked = True Then
                CADENA_NIVELES = CADENA_NIVELES & " AND (MILITAR = 'TRUE')"
            Else
                CADENA_NIVELES = CADENA_NIVELES & " AND (PAME = 'TRUE')"
            End If
            'Situacion del Empleado
            CADENA_NIVELES = CADENA_NIVELES & " AND (ID_SITUACION = 1)"
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
            Dim MiDataTable1 As New DataTable
            Dim MiDataAdapter1 As SqlDataAdapter
            Try
                Conexion.ABD_RRHH()
                CADENAsql4 = "SELECT ORDEN, D_N1, D_N2, D_N3, D_N4, D_N5, D_N6, D_N7, D_N8, D_CARGO_ES, D_GR, CODIGO, AN, ID_M_P FROM [dbo].[VISTA MAESTRO DE CARGOS] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN"
                MiDataAdapter1 = New SqlDataAdapter(CADENAsql4, Conexion.CxRRHH)
                MiDataTable1.Clear()
                MiDataAdapter1.Fill(MiDataTable1)
                If MiDataTable1.Rows.Count > 0 Then
                    For I = 0 To MiDataTable1.Rows.Count - 1
                        CONTADOR_DE_AUSENCIAS = 0
                        If Not IsDBNull(MiDataTable1.Rows(I).Item(0).ToString) Then : v1 = MiDataTable1.Rows(I).Item(0).ToString : Else : v1 = "" : End If
                        If Not IsDBNull(MiDataTable1.Rows(I).Item(1).ToString) Then : v2 = MiDataTable1.Rows(I).Item(1).ToString : Else : v2 = "" : End If
                        If Not IsDBNull(MiDataTable1.Rows(I).Item(2).ToString) Then : v3 = MiDataTable1.Rows(I).Item(2).ToString : Else : v3 = "" : End If
                        If Not IsDBNull(MiDataTable1.Rows(I).Item(3).ToString) Then : v4 = MiDataTable1.Rows(I).Item(3).ToString : Else : v4 = "" : End If
                        If Not IsDBNull(MiDataTable1.Rows(I).Item(4).ToString) Then : v5 = MiDataTable1.Rows(I).Item(4).ToString : Else : v5 = "" : End If
                        If Not IsDBNull(MiDataTable1.Rows(I).Item(5).ToString) Then : v6 = MiDataTable1.Rows(I).Item(5).ToString : Else : v6 = "" : End If
                        If Not IsDBNull(MiDataTable1.Rows(I).Item(6).ToString) Then : v7 = MiDataTable1.Rows(I).Item(6).ToString : Else : v7 = "" : End If
                        If Not IsDBNull(MiDataTable1.Rows(I).Item(7).ToString) Then : v8 = MiDataTable1.Rows(I).Item(7).ToString : Else : v8 = "" : End If
                        If Not IsDBNull(MiDataTable1.Rows(I).Item(8).ToString) Then : v9 = MiDataTable1.Rows(I).Item(8).ToString : Else : v9 = "" : End If
                        If Not IsDBNull(MiDataTable1.Rows(I).Item(9).ToString) Then : v10 = MiDataTable1.Rows(I).Item(9).ToString : Else : v10 = "" : End If
                        If Not IsDBNull(MiDataTable1.Rows(I).Item(10).ToString) Then : v11 = MiDataTable1.Rows(I).Item(10).ToString : Else : v11 = "" : End If
                        If Not IsDBNull(MiDataTable1.Rows(I).Item(11).ToString) Then : v12 = MiDataTable1.Rows(I).Item(11).ToString : Else : v12 = "" : End If
                        If Not IsDBNull(MiDataTable1.Rows(I).Item(12).ToString) Then : v13 = MiDataTable1.Rows(I).Item(12).ToString : Else : v13 = "" : End If
                        Me.Label1.Text = v13
                        If Not IsDBNull(MiDataTable1.Rows(I).Item(13).ToString) Then : v14 = MiDataTable1.Rows(I).Item(13).ToString : Else : v14 = 0 : End If
                        Me.Cursor = Cursors.WaitCursor
                        Me.DateTimePicker3.Value = Me.DateTimePicker1.Value.AddDays(1)
                        Application.DoEvents()
                        Dim FE As Date = Me.DateTimePicker2.Value
                        For I2 As Integer = 1 To DateDiff(DateInterval.Day, Me.DateTimePicker3.Value, FE.AddDays(1))
                            Call BUSCAR_DIA_EN_MAESTRO_DE_SITUACION_DE_PERSONAL()
                            Me.DateTimePicker3.Value = Me.DateTimePicker3.Value.AddDays(1)
                        Next
                        If CONTADOR_DE_AUSENCIAS > 3 Then
                            Call GUARDAR_REGISTRO()
                        End If
                    Next
                    Call CARGAR_DATAGRIDVIEW()
                    Call ARMAR_DATAGRIDVIEW()
                    Me.Cursor = Cursors.Default
                End If
            Catch ex As Exception
                MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
            Finally
                If Conexion.CxRRHH.State = ConnectionState.Open Then
                    Conexion.CBD_RRHH()
                End If
            End Try
        End If
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter("Select * from [TMP CONSULTA AUSENTES DIAS CONTINUOS] WHERE ID_USUARIO = " & isuID_USUARIO & " ORDER BY ORD", Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[TMP CONSULTA AUSENTES DIAS CONTINUOS]")
        DGV.DataSource = MiDataSet.Tables("[TMP CONSULTA AUSENTES DIAS CONTINUOS]")
        Me.Label1.Text = DGV.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Dim D As Date
    Dim FE As Date
    Private Sub ARMAR_DATAGRIDVIEW()
        D = Me.DateTimePicker1.Value.AddDays(1)
        Me.DGV.Columns(0).HeaderText = "Ord"
        Me.DGV.Columns(0).Width = 0
        Me.DGV.Columns(0).Visible = False

        Me.DGV.Columns(1).HeaderText = "Nivel 1"
        Me.DGV.Columns(1).Width = 50
        Me.DGV.Columns(1).Visible = True
        Me.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(2).HeaderText = "Nivel 2"
        Me.DGV.Columns(2).Width = 50
        Me.DGV.Columns(2).Visible = True
        Me.DGV.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(3).HeaderText = "Nivel 3"
        Me.DGV.Columns(3).Width = 50
        Me.DGV.Columns(3).Visible = True
        Me.DGV.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(4).HeaderText = "Nivel 4"
        Me.DGV.Columns(4).Width = 50
        Me.DGV.Columns(4).Visible = True
        Me.DGV.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(5).HeaderText = "Nivel 5"
        Me.DGV.Columns(5).Width = 50
        Me.DGV.Columns(5).Visible = True
        Me.DGV.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(6).HeaderText = "Nivel 6"
        Me.DGV.Columns(6).Width = 50
        Me.DGV.Columns(6).Visible = True
        Me.DGV.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(7).HeaderText = "Nivel 7"
        Me.DGV.Columns(7).Width = 50
        Me.DGV.Columns(7).Visible = True
        Me.DGV.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(8).HeaderText = "Nivel 8"
        Me.DGV.Columns(8).Width = 50
        Me.DGV.Columns(8).Visible = True
        Me.DGV.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(9).HeaderText = "Cargo"
        Me.DGV.Columns(9).Width = 80
        Me.DGV.Columns(9).Visible = True
        Me.DGV.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(10).HeaderText = "Gr"
        Me.DGV.Columns(10).Width = 40
        Me.DGV.Columns(10).Visible = True
        Me.DGV.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(11).HeaderText = "Codigo"
        Me.DGV.Columns(11).Width = 70
        Me.DGV.Columns(11).Visible = True
        Me.DGV.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(12).HeaderText = "Empleado"
        Me.DGV.Columns(12).Width = 110
        Me.DGV.Columns(12).Visible = True
        Me.DGV.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(13).HeaderText = D.Day
        Me.DGV.Columns(13).Width = 30
        Me.DGV.Columns(13).Visible = True
        Me.DGV.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
        Dim D0 As Date = D.AddDays(1)

        Me.DGV.Columns(14).HeaderText = D0.Day
        Me.DGV.Columns(14).Width = 30
        Me.DGV.Columns(14).Visible = True
        Me.DGV.Columns(14).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
        Dim D1 As Date = D0.AddDays(1)

        Me.DGV.Columns(15).HeaderText = D1.Day
        Me.DGV.Columns(15).Width = 30
        Me.DGV.Columns(15).Visible = True
        Me.DGV.Columns(15).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
        Dim D2 As Date = D1.AddDays(1)

        Me.DGV.Columns(16).HeaderText = D2.Day
        Me.DGV.Columns(16).Width = 30
        Me.DGV.Columns(16).Visible = True
        Me.DGV.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
        Dim D3 As Date = D2.AddDays(1)

        Me.DGV.Columns(17).Width = 0
        Me.DGV.Columns(17).Visible = False

    End Sub
    Private Sub GUARDAR_REGISTRO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [TMP CONSULTA AUSENTES DIAS CONTINUOS] (ORD, N1, N2, N3, N4, N5, N6, N7, N8, CARGO, GR, CODIGO, AN, D1, D2, D3, D4, ID_USUARIO)" &
            "values ('" & v1 & "', '" & v2 & "', '" & v3 & "', '" & v4 & "', '" & v5 & "', '" & v6 & "','" & v7 & "','" & v8 & "','" & v9 & "','" & v10 & "','" & v11 & "','" & v12 & "','" & v13 & "', 'FALSE', 'FALSE', 'FALSE', 'FALSE', " & isuID_USUARIO & ")"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub BUSCAR_DIA_EN_MAESTRO_DE_SITUACION_DE_PERSONAL()
        Dim F1 As String = "Convert(VARCHAR(10), Convert(Date, '" + Me.DateTimePicker3.Value + "', 105), 23)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable3 As New DataTable
        Dim MiDataAdapter3 As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            'OMISION DE MARCA/ AUSENCIA INJUSTIFICADA
            MiDataAdapter3 = New SqlDataAdapter("SELECT * FROM [MAESTRO DE SITUACION DE PERSONAL] WHERE (ID_M_P = " & v14 & ") AND (DIA = " & F1 & ") AND (ID_SIT_P = 20 OR ID_SIT_P = 3) ORDER BY ID_M_P, DIA", Conexion.CxRRHH)
            MiDataTable3.Clear()
            MiDataAdapter3.Fill(MiDataTable3)
            If MiDataTable3.Rows.Count <> 0 Then
                EXISTE_DIA_EN_MARCAS = True
                vDIA = True
                CONTADOR_DE_AUSENCIAS = CONTADOR_DE_AUSENCIAS + 1
            Else
                EXISTE_DIA_EN_MARCAS = False
                vDIA = False
                CONTADOR_DE_AUSENCIAS = 0
                'If CONTADOR_DE_AUSENCIAS < 4 Then
                'End If
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ARMAR_CADENA_NIVELES()
        CADENA_NIVELES = ""
        CADENA_NIVELES = "(ID_N1 = " & Me.ComboBox1.SelectedValue & ")"
        If Me.CheckBox1.Checked = True Then
            CADENA_NIVELES = CADENA_NIVELES & " AND (ID_N2 = " & Me.ComboBox2.SelectedValue & ")"
        Else
            GoTo SALIR
        End If
        If Me.CheckBox2.Checked = True Then
            CADENA_NIVELES = CADENA_NIVELES & " AND (ID_N3 = " & Me.ComboBox3.SelectedValue & ")"
        Else
            GoTo SALIR
        End If
        If Me.CheckBox3.Checked = True Then
            CADENA_NIVELES = CADENA_NIVELES & " AND (ID_N4 = " & Me.ComboBox4.SelectedValue & ")"
        Else
            GoTo SALIR
        End If
        If Me.CheckBox4.Checked = True Then
            CADENA_NIVELES = CADENA_NIVELES & " AND (ID_N5 = " & Me.ComboBox5.SelectedValue & ")"
        Else
            GoTo SALIR
        End If
        If Me.CheckBox5.Checked = True Then
            CADENA_NIVELES = CADENA_NIVELES & " AND (ID_N6 = " & Me.ComboBox6.SelectedValue & ")"
        Else
            GoTo SALIR
        End If
        If Me.CheckBox6.Checked = True Then
            CADENA_NIVELES = CADENA_NIVELES & " AND (ID_N7 = " & Me.ComboBox7.SelectedValue & ")"
        Else
            GoTo SALIR
        End If
        If Me.CheckBox7.Checked = True Then
            CADENA_NIVELES = CADENA_NIVELES & " AND (ID_N8 = " & Me.ComboBox8.SelectedValue & ")"
        Else
            GoTo SALIR
        End If
SALIR:
    End Sub
    Private Sub ELIMINAR_TMP_CONSULTA_AUSENTES_DIAS_CONTINUOS()
        Me.Cursor = Cursors.WaitCursor
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "DELETE FROM [TMP CONSULTA AUSENTES DIAS CONTINUOS] WHERE ID_USUARIO = " & isuID_USUARIO & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
        Me.Cursor = Cursors.Default
    End Sub
End Class