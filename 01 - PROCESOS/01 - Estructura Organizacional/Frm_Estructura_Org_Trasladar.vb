Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Estructura_Org_Trasladar
    Private Sub Frm_Estructura_Org_Trasladar_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Estructura_Org_Trasladar_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
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
        'Call ARMAR_CADENA_NIVELES()
        'If isuID_TIPO_USUARIO = 5 Then
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'Else
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'End If
        Call CARGAR_DATAGRIDVIEW_3()
        Call ARMAR_DATAGRIDVIEW_3()
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
        'Call ARMAR_CADENA_NIVELES()
        'If isuID_TIPO_USUARIO = 5 Then
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'Else
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'End If
        'Call CARGAR_DATAGRIDVIEW_3()
        'Call ARMAR_DATAGRIDVIEW_3()
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
        'Call ARMAR_CADENA_NIVELES()
        'If isuID_TIPO_USUARIO = 5 Then
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'Else
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'End If
        'Call CARGAR_DATAGRIDVIEW_3()
        'Call ARMAR_DATAGRIDVIEW_3()
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
        'Call ARMAR_CADENA_NIVELES()
        'If isuID_TIPO_USUARIO = 5 Then
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'Else
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'End If
        'Call CARGAR_DATAGRIDVIEW_3()
        'Call ARMAR_DATAGRIDVIEW_3()
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
        'Call ARMAR_CADENA_NIVELES()
        'If isuID_TIPO_USUARIO = 5 Then
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'Else
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'End If
        'Call CARGAR_DATAGRIDVIEW_3()
        'Call ARMAR_DATAGRIDVIEW_3()
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
        'Call ARMAR_CADENA_NIVELES()
        'If isuID_TIPO_USUARIO = 5 Then
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'Else
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'End If
        'Call CARGAR_DATAGRIDVIEW_3()
        'Call ARMAR_DATAGRIDVIEW_3()
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
        'Call ARMAR_CADENA_NIVELES()
        'If isuID_TIPO_USUARIO = 5 Then
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'Else
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'End If
        'Call CARGAR_DATAGRIDVIEW_3()
        'Call ARMAR_DATAGRIDVIEW_3()
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
        'Call ARMAR_CADENA_NIVELES()
        'If isuID_TIPO_USUARIO = 5 Then
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'Else
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'End If
        'Call CARGAR_DATAGRIDVIEW_3()
        'Call ARMAR_DATAGRIDVIEW_3()
    End Sub
    Private Sub ComboBox8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox8.SelectedIndexChanged
        'Call ARMAR_CADENA_NIVELES()
        'If isuID_TIPO_USUARIO = 5 Then
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'Else
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'End If
        'Call CARGAR_DATAGRIDVIEW_3()
        'Call ARMAR_DATAGRIDVIEW_3()
    End Sub
    Private Sub ARMAR_CADENA_NIVELES()
        CADENA_NIVELES = ""
        CADENA_NIVELES = "(ID_N1 = " & Me.ComboBox1.SelectedValue & ")"
        If Me.CheckBox1.Checked = True Then
            CADENA_NIVELES = CADENA_NIVELES & " AND (ID_N2 = " & Me.ComboBox2.SelectedValue & ")"
        Else
            'CADENA_NIVELES = CADENA_NIVELES & " AND (ID_N2 = NULL)"
            GoTo SALIR
        End If
        If Me.CheckBox2.Checked = True Then
            CADENA_NIVELES = CADENA_NIVELES & " AND (ID_N3 = " & Me.ComboBox3.SelectedValue & ")"
        Else
            'CADENA_NIVELES = CADENA_NIVELES & " AND (ID_N3 = NULL)"
            GoTo SALIR
        End If
        If Me.CheckBox3.Checked = True Then
            CADENA_NIVELES = CADENA_NIVELES & " AND (ID_N4 = " & Me.ComboBox4.SelectedValue & ")"
        Else
            'CADENA_NIVELES = CADENA_NIVELES & " AND (ID_N4 = NULL)"
            GoTo SALIR
        End If
        If Me.CheckBox4.Checked = True Then
            CADENA_NIVELES = CADENA_NIVELES & " AND (ID_N5 = " & Me.ComboBox5.SelectedValue & ")"
        Else
            'CADENA_NIVELES = CADENA_NIVELES & " AND (ID_N5 = NULL)"
            GoTo SALIR
        End If
        If Me.CheckBox5.Checked = True Then
            CADENA_NIVELES = CADENA_NIVELES & " AND (ID_N6 = " & Me.ComboBox6.SelectedValue & ")"
        Else
            'CADENA_NIVELES = CADENA_NIVELES & " AND (ID_N6 = NULL)"
            GoTo SALIR
        End If
        If Me.CheckBox6.Checked = True Then
            CADENA_NIVELES = CADENA_NIVELES & " AND (ID_N7 = " & Me.ComboBox7.SelectedValue & ")"
        Else
            'CADENA_NIVELES = CADENA_NIVELES & " AND (ID_N7 = NULL)"
            GoTo SALIR
        End If
        If Me.CheckBox7.Checked = True Then
            CADENA_NIVELES = CADENA_NIVELES & " AND (ID_N8 = " & Me.ComboBox8.SelectedValue & ")"
        Else
            'CADENA_NIVELES = CADENA_NIVELES & " AND (ID_N8 = NULL)"
            GoTo SALIR
        End If
SALIR:
        'Me.Text = ""
        'Me.Text = CADENA_NIVELES
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If Me.CheckBox1.Checked = True Then
            Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_2()
            NIVEL2sql = Me.ComboBox2.SelectedValue
            'Call ARMAR_CADENA_NIVELES()
        Else
            Me.ComboBox2.DataSource = Nothing
            Me.ComboBox2.ForeColor = Color.Red
            Me.ComboBox2.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
            Me.ComboBox2.Text = "-- SIN DATOS -- "
            NIVEL2sql = Nothing
            Me.CheckBox2.Checked = False
            Me.CheckBox3.Checked = False
            Me.CheckBox4.Checked = False
            'Call ARMAR_CADENA_NIVELES()
        End If
        'If isuID_TIPO_USUARIO = 5 Then
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'Else
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'End If
        'Call CARGAR_DATAGRIDVIEW_3()
        'Call ARMAR_DATAGRIDVIEW_3()
    End Sub
    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If Me.CheckBox2.Checked = True Then
            Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_3()
            NIVEL3sql = Me.ComboBox3.SelectedValue
            'Call ARMAR_CADENA_NIVELES()
        Else
            Me.ComboBox3.DataSource = Nothing
            Me.ComboBox3.ForeColor = Color.Red
            Me.ComboBox3.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
            Me.ComboBox3.Text = "-- SIN DATOS -- "
            NIVEL3sql = Nothing
            Me.CheckBox3.Checked = False
            Me.CheckBox4.Checked = False
            'Call ARMAR_CADENA_NIVELES()
        End If
        'If isuID_TIPO_USUARIO = 5 Then
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'Else
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'End If
        'Call CARGAR_DATAGRIDVIEW_3()
        'Call ARMAR_DATAGRIDVIEW_3()
    End Sub
    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If Me.CheckBox3.Checked = True Then
            Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_4()
            NIVEL4sql = Me.ComboBox4.SelectedValue
            'Call ARMAR_CADENA_NIVELES()
        Else
            Me.ComboBox4.DataSource = Nothing
            Me.ComboBox4.ForeColor = Color.Red
            Me.ComboBox4.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
            Me.ComboBox4.Text = "-- SIN DATOS -- "
            NIVEL4sql = Nothing
            Me.CheckBox4.Checked = False
            'Call ARMAR_CADENA_NIVELES()
        End If
        'If isuID_TIPO_USUARIO = 5 Then
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'Else
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'End If
        'Call CARGAR_DATAGRIDVIEW_3()
        'Call ARMAR_DATAGRIDVIEW_3()
    End Sub
    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If Me.CheckBox4.Checked = True Then
            Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_5()
            NIVEL5sql = Me.ComboBox5.SelectedValue
            'Call ARMAR_CADENA_NIVELES()
        Else
            Me.ComboBox5.DataSource = Nothing
            Me.ComboBox5.ForeColor = Color.Red
            Me.ComboBox5.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
            Me.ComboBox5.Text = "-- SIN DATOS -- "
            NIVEL5sql = Nothing
            'Call ARMAR_CADENA_NIVELES()
        End If
        'If isuID_TIPO_USUARIO = 5 Then
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'Else
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'End If
        'Call CARGAR_DATAGRIDVIEW_3()
        'Call ARMAR_DATAGRIDVIEW_3()
    End Sub
    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If Me.CheckBox5.Checked = True Then
            Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_6()
            NIVEL6sql = Me.ComboBox6.SelectedValue
            'Call ARMAR_CADENA_NIVELES()
        Else
            Me.ComboBox6.DataSource = Nothing
            Me.ComboBox6.ForeColor = Color.Red
            Me.ComboBox6.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
            Me.ComboBox6.Text = "-- SIN DATOS -- "
            NIVEL6sql = Nothing
            'Call ARMAR_CADENA_NIVELES()
        End If
        'If isuID_TIPO_USUARIO = 5 Then
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'Else
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'End If
        'Call CARGAR_DATAGRIDVIEW_3()
        'Call ARMAR_DATAGRIDVIEW_3()
    End Sub
    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        If Me.CheckBox6.Checked = True Then
            Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_7()
            NIVEL7sql = Me.ComboBox7.SelectedValue
            'Call ARMAR_CADENA_NIVELES()
        Else
            Me.ComboBox7.DataSource = Nothing
            Me.ComboBox7.ForeColor = Color.Red
            Me.ComboBox7.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
            Me.ComboBox7.Text = "-- SIN DATOS -- "
            NIVEL7sql = Nothing
            'call ARMAR_CADENA_NIVELES()
        End If
        'If isuID_TIPO_USUARIO = 5 Then
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'Else
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'End If
        'Call CARGAR_DATAGRIDVIEW_3()
        'Call ARMAR_DATAGRIDVIEW_3()
    End Sub
    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        If Me.CheckBox7.Checked = True Then
            Call CARGAR_COMBO_CAT_DE_ESTRUCTURA_NIVEL_8()
            NIVEL8sql = Me.ComboBox8.SelectedValue
            'Call ARMAR_CADENA_NIVELES()
        Else
            Me.ComboBox8.DataSource = Nothing
            Me.ComboBox8.ForeColor = Color.Red
            Me.ComboBox8.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
            Me.ComboBox8.Text = "-- SIN DATOS -- "
            NIVEL8sql = Nothing
            'Call ARMAR_CADENA_NIVELES()
        End If
        'If isuID_TIPO_USUARIO = 5 Then
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'Else
        '    CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN_N1, ORDEN_N2, ORDEN_N3, ORDEN_N4, ORDEN_N5, ORDEN_N6, ORDEN_N7, ORDEN_N8, ORDEN"
        'End If
        'Call CARGAR_DATAGRIDVIEW_3()
        'Call ARMAR_DATAGRIDVIEW_3()
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
    Private Sub CARGAR_DATAGRIDVIEW_3()
        If isuID_TIPO_USUARIO = 5 Then
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " AND D_ADMIN = '" & isuUSUARIO & "' ORDER BY ORDEN"
        Else
            CADENAsql3 = "Select * from [VISTA MAESTRO DE CARGOS SIT] WHERE " & CADENA_NIVELES & " ORDER BY ORDEN"
        End If
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql3, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE CARGOS SIT]")
        DGV3.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE CARGOS SIT]")
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_3()
        Me.DGV3.Columns(0).HeaderText = "Id"
        Me.DGV3.Columns(0).Width = 20
        Me.DGV3.Columns(0).Visible = True
        Me.DGV3.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV3.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Me.DGV3.Columns(1).HeaderText = "ORDEN"
        Me.DGV3.Columns(1).Width = 0
        Me.DGV3.Columns(1).Visible = False

        Me.DGV3.Columns(2).HeaderText = "ID_N1"
        Me.DGV3.Columns(2).Width = 0
        Me.DGV3.Columns(2).Visible = False

        Me.DGV3.Columns(3).HeaderText = "ORDEN_N1"
        Me.DGV3.Columns(3).Width = 0
        Me.DGV3.Columns(3).Visible = False

        Me.DGV3.Columns(4).HeaderText = "Nivel 1"
        Me.DGV3.Columns(4).Width = 0
        Me.DGV3.Columns(4).Visible = False
        Me.DGV3.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(5).HeaderText = "ID_N2"
        Me.DGV3.Columns(5).Width = 0
        Me.DGV3.Columns(5).Visible = False

        Me.DGV3.Columns(6).HeaderText = "ORDEN_N2"
        Me.DGV3.Columns(6).Width = 0
        Me.DGV3.Columns(6).Visible = False

        Me.DGV3.Columns(7).HeaderText = "Nivel 2"
        Me.DGV3.Columns(7).Width = 0
        Me.DGV3.Columns(7).Visible = False
        Me.DGV3.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(8).HeaderText = "ID_N3"
        Me.DGV3.Columns(8).Width = 0
        Me.DGV3.Columns(8).Visible = False

        Me.DGV3.Columns(9).HeaderText = "ORDEN_N3"
        Me.DGV3.Columns(9).Width = 0
        Me.DGV3.Columns(9).Visible = False

        Me.DGV3.Columns(10).HeaderText = "Nivel 3"
        Me.DGV3.Columns(10).Width = 0
        Me.DGV3.Columns(10).Visible = False
        Me.DGV3.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(11).HeaderText = "ID_N4"
        Me.DGV3.Columns(11).Width = 0
        Me.DGV3.Columns(11).Visible = False

        Me.DGV3.Columns(12).HeaderText = "ORDEN_N4"
        Me.DGV3.Columns(12).Width = 0
        Me.DGV3.Columns(12).Visible = False

        Me.DGV3.Columns(13).HeaderText = "Nivel 4"
        Me.DGV3.Columns(13).Width = 0
        Me.DGV3.Columns(13).Visible = False
        Me.DGV3.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(14).HeaderText = "ID_N5"
        Me.DGV3.Columns(14).Width = 0
        Me.DGV3.Columns(14).Visible = False

        Me.DGV3.Columns(15).HeaderText = "ORDEN_N5"
        Me.DGV3.Columns(15).Width = 0
        Me.DGV3.Columns(15).Visible = False

        Me.DGV3.Columns(16).HeaderText = "Nivel 5"
        Me.DGV3.Columns(16).Width = 0
        Me.DGV3.Columns(16).Visible = False
        Me.DGV3.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(17).HeaderText = "ID_N6"
        Me.DGV3.Columns(17).Width = 0
        Me.DGV3.Columns(17).Visible = False

        Me.DGV3.Columns(18).HeaderText = "ORDEN_N6"
        Me.DGV3.Columns(18).Width = 0
        Me.DGV3.Columns(18).Visible = False

        Me.DGV3.Columns(19).HeaderText = "Nivel 6"
        Me.DGV3.Columns(19).Width = 0
        Me.DGV3.Columns(19).Visible = False
        Me.DGV3.Columns(19).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(20).HeaderText = "ID_N7"
        Me.DGV3.Columns(20).Width = 0
        Me.DGV3.Columns(20).Visible = False

        Me.DGV3.Columns(21).HeaderText = "ORDEN_N7"
        Me.DGV3.Columns(21).Width = 0
        Me.DGV3.Columns(21).Visible = False

        Me.DGV3.Columns(22).HeaderText = "Nivel 7"
        Me.DGV3.Columns(22).Width = 0
        Me.DGV3.Columns(22).Visible = False
        Me.DGV3.Columns(22).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(23).HeaderText = "ID_N8"
        Me.DGV3.Columns(23).Width = 0
        Me.DGV3.Columns(23).Visible = False

        Me.DGV3.Columns(24).HeaderText = "ORDEN_N8"
        Me.DGV3.Columns(24).Width = 0
        Me.DGV3.Columns(24).Visible = False

        Me.DGV3.Columns(25).HeaderText = "Nivel 8"
        Me.DGV3.Columns(25).Width = 0
        Me.DGV3.Columns(25).Visible = False
        Me.DGV3.Columns(25).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(26).HeaderText = "ID_T_ES"
        Me.DGV3.Columns(26).Width = 0
        Me.DGV3.Columns(26).Visible = False

        Me.DGV3.Columns(27).HeaderText = "D_T_ES"
        Me.DGV3.Columns(27).Width = 0
        Me.DGV3.Columns(27).Visible = False

        Me.DGV3.Columns(28).HeaderText = "No. Orden [MIXTO]"
        Me.DGV3.Columns(28).Width = 0
        Me.DGV3.Columns(28).Visible = False
        Me.DGV3.Columns(28).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV3.Columns(28).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV3.Columns(28).DefaultCellStyle.BackColor = Color.Khaki

        Me.DGV3.Columns(29).HeaderText = "No. Orden [MILITAR]"
        Me.DGV3.Columns(29).Width = 0
        Me.DGV3.Columns(29).Visible = False
        Me.DGV3.Columns(29).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV3.Columns(29).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV3.Columns(29).DefaultCellStyle.BackColor = Color.LightGreen

        Me.DGV3.Columns(30).HeaderText = "No. Orden [PAME]"
        Me.DGV3.Columns(30).Width = 0
        Me.DGV3.Columns(30).Visible = False
        Me.DGV3.Columns(30).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV3.Columns(30).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV3.Columns(30).DefaultCellStyle.BackColor = Color.LightBlue

        Me.DGV3.Columns(31).HeaderText = "ID_CARGO_ES"
        Me.DGV3.Columns(31).Width = 0
        Me.DGV3.Columns(31).Visible = False

        Me.DGV3.Columns(32).HeaderText = "Cargo"
        Me.DGV3.Columns(32).Width = 190
        Me.DGV3.Columns(32).Visible = True
        Me.DGV3.Columns(32).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(33).HeaderText = "ID_GP"
        Me.DGV3.Columns(33).Width = 0
        Me.DGV3.Columns(33).Visible = False

        Me.DGV3.Columns(34).HeaderText = "Grado Plantilla"
        Me.DGV3.Columns(34).Width = 0
        Me.DGV3.Columns(34).Visible = False
        Me.DGV3.Columns(34).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(35).HeaderText = "ID_GR"
        Me.DGV3.Columns(35).Width = 0
        Me.DGV3.Columns(35).Visible = False

        Me.DGV3.Columns(36).HeaderText = "Grado Real"
        Me.DGV3.Columns(36).Width = 0
        Me.DGV3.Columns(36).Visible = False
        Me.DGV3.Columns(36).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(37).HeaderText = "ID_CAT_SALARIAL"
        Me.DGV3.Columns(37).Width = 0
        Me.DGV3.Columns(37).Visible = False

        Me.DGV3.Columns(38).HeaderText = "D_CAT_SALARIAL"
        Me.DGV3.Columns(38).Width = 0
        Me.DGV3.Columns(38).Visible = False

        Me.DGV3.Columns(39).HeaderText = "ID_M_P"
        Me.DGV3.Columns(39).Width = 0
        Me.DGV3.Columns(39).Visible = False

        Me.DGV3.Columns(40).HeaderText = "Codigo"
        Me.DGV3.Columns(40).Width = 80
        Me.DGV3.Columns(40).Visible = True
        Me.DGV3.Columns(40).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(41).HeaderText = "Apellidos y Nombres"
        Me.DGV3.Columns(41).Width = 200
        Me.DGV3.Columns(41).Visible = True
        Me.DGV3.Columns(41).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV3.Columns(41).DefaultCellStyle.BackColor = Color.Khaki

        Me.DGV3.Columns(42).HeaderText = "Nombres y Apellidos"
        Me.DGV3.Columns(42).Width = 0
        Me.DGV3.Columns(42).Visible = False

        Me.DGV3.Columns(43).HeaderText = "ID_SITUACION"
        Me.DGV3.Columns(43).Width = 0
        Me.DGV3.Columns(43).Visible = False

        Me.DGV3.Columns(44).HeaderText = "Situacion"
        Me.DGV3.Columns(44).Width = 60
        Me.DGV3.Columns(44).Visible = True
        Me.DGV3.Columns(44).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(45).HeaderText = "Jefe"
        Me.DGV3.Columns(45).Width = 0
        Me.DGV3.Columns(45).Visible = False
        Me.DGV3.Columns(45).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(46).HeaderText = "Mixta"
        Me.DGV3.Columns(46).Width = 0
        Me.DGV3.Columns(46).Visible = False
        Me.DGV3.Columns(46).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(47).HeaderText = "Militar"
        Me.DGV3.Columns(47).Width = 0
        Me.DGV3.Columns(47).Visible = False
        Me.DGV3.Columns(47).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(48).HeaderText = "Pame"
        Me.DGV3.Columns(48).Width = 0
        Me.DGV3.Columns(48).Visible = False
        Me.DGV3.Columns(48).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(49).HeaderText = "ID_ESTABLECIMIENTO"
        Me.DGV3.Columns(49).Width = 0
        Me.DGV3.Columns(49).Visible = False

        Me.DGV3.Columns(50).HeaderText = "Establecimiento"
        Me.DGV3.Columns(50).Width = 0
        Me.DGV3.Columns(50).Visible = False
        Me.DGV3.Columns(50).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(51).HeaderText = "D_ADMIN"
        Me.DGV3.Columns(51).Width = 0
        Me.DGV3.Columns(51).Visible = False
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.DGV3.RowCount = 0 Then
            MsgBox("No existen cargos que Trasladar", vbInformation, "Mensaje del Sistema")
            Me.DGV3.Focus()
            Exit Sub
        End If
        Call RECORRER_DGV_CARGOS()
    End Sub
    Private Sub RECORRER_DGV_CARGOS()
        If Me.DGV3.RowCount <> 0 Then
            For I = 0 To Me.DGV3.RowCount - 1
                If DGV3.Rows(I).Selected = True Then
                    TvmcsID_M_C = Me.DGV3.Item(0, I).Value
                    Call EDITAR_REGISTRO()
                End If
            Next
            Call CARGAR_DATAGRIDVIEW_3()
            Call ARMAR_DATAGRIDVIEW_3()
        Else
        End If
    End Sub
    Dim vN11, vN22, vN33, vN44, vN55, vN66, vN77, vN88 As Integer
    Dim vN1, vN2, vN3, vN4, vN5, vN6, vN7, vN8 As String
    Private Sub EDITAR_REGISTRO()
        'NIVEL 1
        If Me.ComboBox1.SelectedValue = 0 Then
            vN11 = Nothing
        Else
            vN11 = Me.ComboBox1.SelectedValue
        End If
        'NIVEL 2
        If Me.ComboBox2.SelectedValue = 0 Then
            vN22 = Nothing
        Else
            vN22 = Me.ComboBox2.SelectedValue
        End If
        'NIVEL 3
        If Me.ComboBox3.SelectedValue = 0 Then
            vN33 = Nothing
        Else
            vN33 = Me.ComboBox3.SelectedValue
        End If
        'NIVEL 4
        If Me.ComboBox4.SelectedValue = 0 Then
            vN44 = Nothing
        Else
            vN44 = Me.ComboBox4.SelectedValue
        End If
        'NIVEL 5
        If Me.ComboBox5.SelectedValue = 0 Then
            vN55 = Nothing
        Else
            vN55 = Me.ComboBox5.SelectedValue
        End If
        'NIVEL 6
        If Me.ComboBox6.SelectedValue = 0 Then
            vN66 = Nothing
        Else
            vN66 = Me.ComboBox6.SelectedValue
        End If
        'NIVEL 7
        If Me.ComboBox7.SelectedValue = 0 Then
            vN77 = Nothing
        Else
            vN77 = Me.ComboBox7.SelectedValue
        End If
        'NIVEL 8
        If Me.ComboBox8.SelectedValue = 0 Then
            vN88 = Nothing
        Else
            vN88 = Me.ComboBox8.SelectedValue
        End If

        If vN11 = 0 Then : vN1 = "NULL" : Else : vN1 = Me.ComboBox1.SelectedValue : End If
        If vN22 = 0 Then : vN2 = "NULL" : Else : vN2 = Me.ComboBox2.SelectedValue : End If
        If vN33 = 0 Then : vN3 = "NULL" : Else : vN3 = Me.ComboBox3.SelectedValue : End If
        If vN44 = 0 Then : vN4 = "NULL" : Else : vN4 = Me.ComboBox4.SelectedValue : End If
        If vN55 = 0 Then : vN5 = "NULL" : Else : vN5 = Me.ComboBox5.SelectedValue : End If
        If vN66 = 0 Then : vN6 = "NULL" : Else : vN6 = Me.ComboBox6.SelectedValue : End If
        If vN77 = 0 Then : vN7 = "NULL" : Else : vN7 = Me.ComboBox7.SelectedValue : End If
        If vN88 = 0 Then : vN8 = "NULL" : Else : vN8 = Me.ComboBox8.SelectedValue : End If

        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "UPDATE [MAESTRO DE CARGOS] SET ID_N1 = " & vN1 & ",  ID_N2 = " & vN2 & ",  ID_N3 = " & vN3 & ",  ID_N4 = " & vN4 & ",  ID_N5 = " & vN5 & ", ID_N6 = " & vN6 & ", ID_N7 = " & vN7 & ", ID_N8 = " & vN8 & " WHERE ID_M_C = " & CInt(TvmcsID_M_C) & ""
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
End Class