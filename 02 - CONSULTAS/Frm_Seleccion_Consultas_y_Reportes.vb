Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient

Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class Frm_Seleccion_Consultas_y_Reportes
    Private Sub Frm_Seleccion_Consultas_y_Reportes_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            NOMBRE_REPORTE = ""
            INFORMES = False
            Frm_Principal.Timer1.Start()
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Seleccion_Consultas_y_Reportes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Frm_Principal.Timer1.Stop()
        NOMBRE_REPORTE = ""
        INFORMES = True
        Dim tp As TabPage = TabControl.TabPages(0)
        TabControl.SelectedTab = tp
        Call SELECCIONAR_CADENA()
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
        Me.RadioButton3.Checked = True
        Me.RadioButton8.Checked = True
        Call CARGAR_COMBO_CAT_DE_GRADO_REAL()
        Call CARGAR_COMBO_CAT_DE_GRADO_PLANTILLAL()
        Call CARGAR_COMBO_CAT_DE_SITUACION()
        Me.ComboBox11.Text = "OCUPADO"
        Call CARGAR_COMBO_CAT_DE_CARGOS_DE_ESTRUCTURA()
        Call CARGAR_COMBO_CAT_DE_ESTABLECIMIENTO()
        RemoveHandler Me.ComboBox14.SelectedIndexChanged, AddressOf ComboBox14_SelectedIndexChanged
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_MOVIMIENTO_DE_EXP()
        Call CARGAR_COMBO_CAT_DE_SUB_TIPO_DE_MOVIMIENTO_REAL_DE_EXP()
        Call CARGAR_COMBO_CAT_DE_SUB_TIPO_DE_MOVIMIENTO_LEGAL_DE_EXP()
        AddHandler Me.ComboBox14.SelectedIndexChanged, AddressOf ComboBox14_SelectedIndexChanged
        RemoveHandler Me.ComboBox17.SelectedIndexChanged, AddressOf ComboBox17_SelectedIndexChanged
        Call CARGAR_COMBO_SOL_CAT_DE_PROCESOS_PRINCIPALES()
        Call CARGAR_COMBO_SOL_CAT_DE_RESULTADOS()
        AddHandler Me.ComboBox17.SelectedIndexChanged, AddressOf ComboBox17_SelectedIndexChanged
        Call CARGAR_COMBO_CAT_DE_FUNCIONALIDAD()
        Call CARGAR_COMBO_CAT_DE_CAUSAS_REALES_DE_EVENTOS()
        Call CARGAR_COMBO_CAT_DE_EVENTUALIDADES()
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_ESTUDIO()
        Call CARGAR_COMBO_CAT_DE_NIVEL_DE_COMPETENCIA()
        Call CARGAR_COMBO_CAT_DE_RELOJ()
        Call CARGAR_COMBO_CAT_DE_CLASIFICACION_DE_PERSONAL()
        Call CARGAR_COMBO_CAT_DE_FORMATOS_DE_SALIDA()
        RemoveHandler Me.ComboBox27.SelectedIndexChanged, AddressOf ComboBox27_SelectedIndexChanged
        Call CARGAR_COMBO_CAT_DE_CLASIFICACION_DE_CONECORACIONES()
        Call CARGAR_COMBO_CAT_DE_CONDECORACIONES()
        AddHandler Me.ComboBox27.SelectedIndexChanged, AddressOf ComboBox27_SelectedIndexChanged
        RemoveHandler Me.ComboBox29.SelectedIndexChanged, AddressOf ComboBox29_SelectedIndexChanged
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_SITUACION_DE_PERSONAL()
        Me.ComboBox29.Text = "AUSENTE"
        Call CARGAR_COMBO_CAT_DE_SITUACION_DE_PERSONAL()
        Me.ComboBox30.Text = "VACACIONES"
        AddHandler Me.ComboBox29.SelectedIndexChanged, AddressOf ComboBox29_SelectedIndexChanged
        Me.ComboBox25.Text = "ENTRADA"
        Me.TextBox1.Text = ""
        Call DESACTIVAR_TODO()
    End Sub
    Private Sub DESACTIVAR_TODO()
        '--------------- NIVELES ORGANIZACIONALES
        Me.ComboBox1.Enabled = False        'NIVEL 1
        Me.ComboBox2.Enabled = False        'NIVEL 2
        Me.CheckBox1.Enabled = False        'CHECK N2
        Me.ComboBox3.Enabled = False        'NIVEL 3
        Me.CheckBox2.Enabled = False        'CHECK N3
        Me.ComboBox4.Enabled = False        'NIVEL 4
        Me.CheckBox3.Enabled = False        'CHECK N4
        Me.ComboBox5.Enabled = False        'NIVEL 5
        Me.CheckBox4.Enabled = False        'CHECK N5
        Me.ComboBox6.Enabled = False        'NIVEL 6
        Me.CheckBox5.Enabled = False        'CHECK N6
        Me.ComboBox7.Enabled = False        'NIVEL 7
        Me.CheckBox6.Enabled = False        'CHECK N7
        Me.ComboBox8.Enabled = False        'NIVEL 8
        Me.CheckBox7.Enabled = False        'CHECK N8
        Me.Button1.Enabled = False          'BOTON NIVELES ORGANIZACIONALES
        '--------------- TIPO DE ESTRUCTURAS
        Me.RadioButton1.Enabled = False     'ESTRUCTURA MIXTA
        Me.RadioButton2.Enabled = False     'ESTRUCTURA MILITAR
        Me.RadioButton3.Enabled = False     'ESTRUCTURA PAME
        Me.RadioButton4.Enabled = False     'JEFE
        Me.Button2.Enabled = False          'BOTON TIPO DE ESTRUCTURAS
        '--------------- GRADOS, CARGOS, OTROS
        Me.ComboBox9.Enabled = False        'GRADO REAL
        Me.Button3.Enabled = False          'BOTON GR
        Me.ComboBox10.Enabled = False       'GRADO PLANTILLA
        Me.Button4.Enabled = False          'BOTON GP
        Me.ComboBox11.Enabled = False       'SITUACION DEL CARGO
        Me.Button5.Enabled = False          'BOTON SITUACION CARGO
        Me.ComboBox12.Enabled = False       'CARGO
        Me.Button6.Enabled = False          'BOTON CARGO
        Me.ComboBox13.Enabled = False       'ESTABLECIMIENTO
        Me.Button7.Enabled = False          'BOTON ESTABLECIMIENTO
        '--------------- MOVIMIENTOS
        Me.ComboBox14.Enabled = False       'TIPO DE MOVIMIENTO
        Me.Button8.Enabled = False          'BOTON TIPO MOV
        Me.ComboBox15.Enabled = False       'CAUSAL REAL 
        Me.Button9.Enabled = False          'BOTON CAUSAL REAL
        Me.ComboBox16.Enabled = False       'CAUSAL LEGAL
        Me.Button10.Enabled = False         'BOTON CAUSAL LEGAL
        '--------------- FECHAS
        Me.RadioButton5.Enabled = False     'FECHA DE INGRESO PAME
        Me.RadioButton6.Enabled = False     'FECHA DE INGRESO EN
        Me.RadioButton7.Enabled = False     'FECHA DE NACIMIENTO
        Me.RadioButton8.Enabled = False     'FECHA DE SITUACION
        Me.RadioButton9.Enabled = False     'FECHA DE MARCAS
        Me.RadioButton10.Enabled = False    'FECHA DE PROCESO DE ASPIRANTE
        Me.RadioButton11.Enabled = False    'FECHA DE MOVIMIENTOS
        Me.RadioButton15.Enabled = False    'FECHA DE EVENTUALIDADES
        Me.RadioButton16.Enabled = False    'FECHA DE SUBSIDIOS
        Me.DateTimePicker1.Enabled = False  'DESE
        Me.DateTimePicker2.Enabled = False  'HASTA
        Me.Button11.Enabled = False         'BOTON FECHAS
        '--------------- ASPIRANTES
        Me.ComboBox17.Enabled = False       'PROCESO
        Me.Button12.Enabled = False         'BOTON  PROCESO
        Me.ComboBox18.Enabled = False       'RESULTADO
        Me.Button13.Enabled = False         'BOTON RESULTADO
        '--------------- OTROS FILTROS
        Me.ComboBox19.Enabled = False       'FUNCIONALIDAD DEL EMPLEADO
        Me.Button14.Enabled = False         'BOTON FUNCIONALIDAD DEL EMPLEADO
        Me.ComboBox20.Enabled = False       'CAUSA REAL DE EVENTUALIDAD
        Me.Button15.Enabled = False         'BOTON CAUSA REAL DE EVENTUALIDAD
        Me.ComboBox21.Enabled = False       'TIPO DE EVENTUALIDAD
        Me.Button16.Enabled = False         'BOTON TIPO DE EVENTUALIDAD
        Me.ComboBox22.Enabled = False       'ESTUDIOS
        Me.Button17.Enabled = False         'BOTON ESTUDIOS
        Me.ComboBox23.Enabled = False       'NIVEL DE COMPETENCIA
        Me.Button18.Enabled = False         'BOTON NIVEL DE COMPETENCIA
        Me.ComboBox24.Enabled = False       'RELOJ DE MARCACION
        Me.Button19.Enabled = False         'BOTON RELOJ DE MARCACION
        Me.ComboBox25.Enabled = False       'TIPO DE MARCA
        Me.Button20.Enabled = False         'BOTON TIPO DE MARCA
        Me.ComboBox26.Enabled = False       'CLASIFICACION DEL EMPLEADO
        Me.Button21.Enabled = False         'BOTON CLASIFICACION
        Me.TextBox1.Enabled = False         'CODIGO DEL EMPLEADO
        Me.Button22.Enabled = False         'BOTON CODIGO EMPLEADO
        Me.RadioButton12.Enabled = False    'EDAD
        Me.RadioButton13.Enabled = False    'ANTIGUEDAD PAME
        Me.RadioButton14.Enabled = False    'ANTIGUEDAD EN
        Me.TextBox2.Enabled = False         'ANTIGUEDADES A
        Me.TextBox3.Enabled = False         'ANTIGUEDADES B
        Me.Button23.Enabled = False         'BOTON ANTIGUEDADES
        '--------------- CONDECORACIONES
        Me.ComboBox27.Enabled = False       'CLASIFICACION
        Me.Button25.Enabled = False         'BOTON CLASIFICACION
        Me.ComboBox28.Enabled = False       'TIPO
        Me.Button26.Enabled = False         'BOTON TIPO

        Me.TextBox4.Enabled = False         'CONTRATOS
        Me.Button24.Enabled = False         'BOTON CONTRATOS
        '--------------- SITUACIONES
        Me.ComboBox29.Enabled = False       'TIPO SITUACION
        Me.Button27.Enabled = False         'BOTON TIPO SITUACION
        Me.ComboBox30.Enabled = False       'SITUACION
        Me.Button28.Enabled = False         'BOTON SITUACION
        Me.DateTimePicker3.Enabled = False  'DESDE
        Me.DateTimePicker4.Enabled = False  'HASTA
        Me.Button29.Enabled = False         'BOTON FECHAS
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
            'Me.ComboBox2.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
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
            'Me.ComboBox3.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
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
            'Me.ComboBox4.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
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
            'Me.ComboBox5.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
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
            'Me.ComboBox6.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
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
            'Me.ComboBox7.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
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
            'Me.ComboBox8.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
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
                    'Me.ComboBox8.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
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
            'Me.ComboBox8.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
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
                    'Me.ComboBox7.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
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
            'Me.ComboBox7.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
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
                    'Me.ComboBox6.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
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
            'Me.ComboBox6.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
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
                    'Me.ComboBox5.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
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
            'Me.ComboBox5.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
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
                    'Me.ComboBox4.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
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
            'Me.ComboBox4.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
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
                    'Me.ComboBox3.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
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
            'Me.ComboBox3.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
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
                'Me.ComboBox2.Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Italic)
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
    Private Sub CARGAR_COMBO_CAT_DE_GRADO_REAL()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE GRADO REAL] ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                With Me.ComboBox9
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_GR"
                End With
            Else
                Me.ComboBox9.DataSource = Nothing
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_GRADO_PLANTILLAL()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE GRADO PLANTILLA] ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                With Me.ComboBox10
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_GP"
                End With
            Else
                Me.ComboBox10.DataSource = Nothing
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_SITUACION()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE SITUACION] ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                With Me.ComboBox11
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_SITUACION"
                End With
            Else
                Me.ComboBox11.DataSource = Nothing
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_CARGOS_DE_ESTRUCTURA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE CARGOS DE ESTRUCTURA] ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                With Me.ComboBox12
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_CARGO_ES"
                End With
            Else
                Me.ComboBox12.DataSource = Nothing
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTABLECIMIENTO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTABLECIMIENTO] ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                With Me.ComboBox13
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_ESTABLECIMIENTO"
                End With
            Else
                Me.ComboBox13.DataSource = Nothing
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_MOVIMIENTO_DE_EXP()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE MOVIMIENTO DE EXP] ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                With Me.ComboBox14
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_T_M_EXP"
                End With
            Else
                Me.ComboBox14.DataSource = Nothing
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_SUB_TIPO_DE_MOVIMIENTO_REAL_DE_EXP()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE SUB TIPO DE MOVIMIENTO REAL DE EXP] WHERE ID_T_M_EXP = " & Me.ComboBox14.SelectedValue & " ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                With Me.ComboBox15
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_ST_M_R_EXP"
                End With
            Else
                Me.ComboBox15.DataSource = Nothing
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_SUB_TIPO_DE_MOVIMIENTO_LEGAL_DE_EXP()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE SUB TIPO DE MOVIMIENTO LEGAL DE EXP] WHERE ID_T_M_EXP = " & Me.ComboBox14.SelectedValue & "ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                With Me.ComboBox16
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_ST_M_L_EXP"
                End With
            Else
                Me.ComboBox16.DataSource = Nothing
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_SOL_CAT_DE_PROCESOS_PRINCIPALES()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [SOL_CAT DE PROCESOS PRINCIPALES] ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                With Me.ComboBox17
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_PROC"
                End With
            Else
                Me.ComboBox17.DataSource = Nothing
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_SOL_CAT_DE_RESULTADOS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [SOL_CAT DE RESULTADOS] WHERE ID_PROC = " & Me.ComboBox17.SelectedValue & " ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                With Me.ComboBox18
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_R_SOL"
                End With
            Else
                Me.ComboBox18.DataSource = Nothing
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_FUNCIONALIDAD()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE FUNCIONALIDAD] ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                With Me.ComboBox19
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_F"
                End With
            Else
                Me.ComboBox19.DataSource = Nothing
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_CAUSAS_REALES_DE_EVENTOS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE CAUSAS REALES DE EVENTOS] ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                With Me.ComboBox20
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_CR_EVENTOS"
                End With
            Else
                Me.ComboBox20.DataSource = Nothing
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_EVENTUALIDADES()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE EVENTUALIDADES] ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                With Me.ComboBox21
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_E"
                End With
            Else
                Me.ComboBox21.DataSource = Nothing
            End If
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
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE ESTUDIO] ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                With Me.ComboBox22
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_T_ESTUDIO"
                End With
            Else
                Me.ComboBox22.DataSource = Nothing
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_NIVEL_DE_COMPETENCIA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE NIVEL DE COMPETENCIA] ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                With Me.ComboBox23
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_N_COMPETENCIA"
                End With
            Else
                Me.ComboBox23.DataSource = Nothing
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_RELOJ()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE RELOJ] ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                With Me.ComboBox24
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "IDRELOJ"
                End With
            Else
                Me.ComboBox24.DataSource = Nothing
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_CLASIFICACION_DE_PERSONAL()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE CLASIFICACION DE PERSONAL] ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                With Me.ComboBox26
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_C_PE"
                End With
            Else
                Me.ComboBox26.DataSource = Nothing
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_FORMATOS_DE_SALIDA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE FORMATOS DE SALIDA] ORDER BY ID_F_S", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                With Me.ComboBox100
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_F_S"
                End With
            Else
                Me.ComboBox100.DataSource = Nothing
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_CONDECORACIONES()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE CONDECORACIONES] where ID_C_CONDE = " & Me.ComboBox27.SelectedValue & " ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                With Me.ComboBox28
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_CONDE"
                End With
            Else
                Me.ComboBox100.DataSource = Nothing
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_CLASIFICACION_DE_CONECORACIONES()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE CLASIFICACION DE CONECORACIONES] ORDER BY ID_C_CONDE", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                With Me.ComboBox27
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_C_CONDE"
                End With
            Else
                Me.ComboBox100.DataSource = Nothing
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_SITUACION_DE_PERSONAL()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE SITUACION DE PERSONAL] ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                With Me.ComboBox29
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_T_SIT_P"
                End With
            Else
                Me.ComboBox100.DataSource = Nothing
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_SITUACION_DE_PERSONAL()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE SITUACION DE PERSONAL] where ID_T_SIT_P = " & Me.ComboBox29.SelectedValue & " ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                With Me.ComboBox30
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_SIT_P"
                End With
            Else
                Me.ComboBox100.DataSource = Nothing
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button200_Click(sender As Object, e As EventArgs) Handles Button200.Click
        Me.Lista_Principal.Items.Clear()
        Call LIMPIAR_LISTAS_DE_SELECCION()
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button300.Click
        NOMBRE_REPORTE = ""
        INFORMES = False
        Frm_Principal.Timer1.Start()
        Me.Close()
    End Sub
    Private Sub LIMPIAR_LISTAS()
        Me.ListBox1.Items.Clear()
        Me.ListBox2.Items.Clear()
        Me.ListBox3.Items.Clear()
        Me.ListBox4.Items.Clear()
        Me.ListBox5.Items.Clear()
        Me.ListBox6.Items.Clear()
        Me.ListBox7.Items.Clear()
        Me.ListBox8.Items.Clear()
    End Sub
    Private Sub LIMPIAR_LISTAS_DE_SELECCION()
        Me.Lista101.Items.Clear()
        Me.Lista102.Items.Clear()
        Me.Lista103.Items.Clear()
        Me.Lista104.Items.Clear()
        Me.Lista105.Items.Clear()
        Me.Lista106.Items.Clear()
        Me.Lista107.Items.Clear()
        Me.Lista108.Items.Clear()
        Me.Lista109.Items.Clear()
        Me.Lista110.Items.Clear()
        Me.Lista111.Items.Clear()
        Me.Lista112.Items.Clear()
        Me.Lista113.Items.Clear()
        Me.Lista114.Items.Clear()
        Me.Lista115.Items.Clear()
        Me.Lista116.Items.Clear()
        Me.Lista117.Items.Clear()
        Me.Lista118.Items.Clear()
        Me.Lista119.Items.Clear()
        Me.Lista120.Items.Clear()
        Me.Lista121.Items.Clear()
        Me.Lista122.Items.Clear()
        Me.Lista123.Items.Clear()
        Me.Lista124.Items.Clear()
        Me.Lista125.Items.Clear()
        Me.Lista126.Items.Clear()
        Me.Lista127.Items.Clear()
        Me.Lista128.Items.Clear()
        Me.Lista129.Items.Clear()
        Me.Lista130.Items.Clear()
        Me.Lista131.Items.Clear()
        Me.Lista132.Items.Clear()
        Me.Lista133.Items.Clear()
        Me.Lista134.Items.Clear()
        Me.Lista135.Items.Clear()
        Me.Lista136.Items.Clear()
        Me.Lista137.Items.Clear()
        Me.Lista138.Items.Clear()
        Me.Lista139.Items.Clear()
        Me.Lista140.Items.Clear()
        Me.Lista141.Items.Clear()
        Me.Lista142.Items.Clear()
        Me.Lista143.Items.Clear()
    End Sub
    Private Sub TabControl_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TabControl.Selecting
        Call SELECCIONAR_CADENA()
    End Sub
    Private Sub SELECCIONAR_CADENA()
        If Me.TabControl.TabPages(Me.TabControl.SelectedIndex).Tag = 1 Then     'ORGANICA
            Call LIMPIAR_LISTAS()
            CONTROL_CARGA_LISTA = 1
            CADENAsql = "SELECT * FROM [MAESTRO DE CONSULTAS Y REPORTES] WHERE TIPO = 1 AND ACTIVO = 'TRUE' ORDER BY DESCRIPCION"
            Call CARGAR_LISTA_COLUMNAS()
            Call CONTAR_CANTIDAD_DE_REPORTES()
            GoTo SALIR
        End If
        If Me.TabControl.TabPages(Me.TabControl.SelectedIndex).Tag = 2 Then     'MOVIMIENTOS
            Call LIMPIAR_LISTAS()
            CONTROL_CARGA_LISTA = 2
            CADENAsql = "SELECT * FROM [MAESTRO DE CONSULTAS Y REPORTES] WHERE TIPO = 2 AND ACTIVO = 'TRUE' ORDER BY DESCRIPCION"
            Call CARGAR_LISTA_COLUMNAS()
            Call CONTAR_CANTIDAD_DE_REPORTES()
            GoTo SALIR
        End If
        If Me.TabControl.TabPages(Me.TabControl.SelectedIndex).Tag = 3 Then     'ASPIRANTES
            Call LIMPIAR_LISTAS()
            CONTROL_CARGA_LISTA = 3
            CADENAsql = "SELECT * FROM [MAESTRO DE CONSULTAS Y REPORTES] WHERE TIPO = 3 AND ACTIVO = 'TRUE' ORDER BY DESCRIPCION"
            Call CARGAR_LISTA_COLUMNAS()
            Call CONTAR_CANTIDAD_DE_REPORTES()
            GoTo SALIR
        End If
        If Me.TabControl.TabPages(Me.TabControl.SelectedIndex).Tag = 4 Then     'EXPEDIENTES
            Call LIMPIAR_LISTAS()
            CONTROL_CARGA_LISTA = 4
            CADENAsql = "SELECT * FROM [MAESTRO DE CONSULTAS Y REPORTES] WHERE TIPO = 4 AND ACTIVO = 'TRUE' ORDER BY DESCRIPCION"
            Call CARGAR_LISTA_COLUMNAS()
            Call CONTAR_CANTIDAD_DE_REPORTES()
            GoTo SALIR
        End If
        If Me.TabControl.TabPages(Me.TabControl.SelectedIndex).Tag = 5 Then     'MEDICOS ACREDITADOS
            Call LIMPIAR_LISTAS()
            CONTROL_CARGA_LISTA = 5
            CADENAsql = "SELECT * FROM [MAESTRO DE CONSULTAS Y REPORTES] WHERE TIPO = 5 AND ACTIVO = 'TRUE' ORDER BY DESCRIPCION"
            Call CARGAR_LISTA_COLUMNAS()
            Call CONTAR_CANTIDAD_DE_REPORTES()
            GoTo SALIR
        End If
        If Me.TabControl.TabPages(Me.TabControl.SelectedIndex).Tag = 6 Then     'BECADOS
            Call LIMPIAR_LISTAS()
            CONTROL_CARGA_LISTA = 6
            CADENAsql = "SELECT * FROM [MAESTRO DE CONSULTAS Y REPORTES] WHERE TIPO = 6 AND ACTIVO = 'TRUE' ORDER BY DESCRIPCION"
            Call CARGAR_LISTA_COLUMNAS()
            Call CONTAR_CANTIDAD_DE_REPORTES()
            GoTo SALIR
        End If
        If Me.TabControl.TabPages(Me.TabControl.SelectedIndex).Tag = 7 Then     'ASISTENCIA Y SITUACIONES DEL PERSONAL
            Call LIMPIAR_LISTAS()
            CONTROL_CARGA_LISTA = 7
            CADENAsql = "SELECT * FROM [MAESTRO DE CONSULTAS Y REPORTES] WHERE TIPO = 7 AND ACTIVO = 'TRUE' ORDER BY DESCRIPCION"
            Call CARGAR_LISTA_COLUMNAS()
            Call CONTAR_CANTIDAD_DE_REPORTES()
            GoTo SALIR
        End If
        If Me.TabControl.TabPages(Me.TabControl.SelectedIndex).Tag = 8 Then     'HIGIENE Y SEGURIDAD
            Call LIMPIAR_LISTAS()
            CONTROL_CARGA_LISTA = 8
            CADENAsql = "SELECT * FROM [MAESTRO DE CONSULTAS Y REPORTES] WHERE TIPO = 8 AND ACTIVO = 'TRUE' ORDER BY DESCRIPCION"
            Call CARGAR_LISTA_COLUMNAS()
            Call CONTAR_CANTIDAD_DE_REPORTES()
            GoTo SALIR
        End If
SALIR:
    End Sub
    Private Sub CONTAR_CANTIDAD_DE_REPORTES()
        If CONTROL_CARGA_LISTA = 1 Then
            Me.Label1.Text = Me.ListBox1.Items.Count & " Reporte(s)"
            GoTo LEE_OTRO
        End If
        If CONTROL_CARGA_LISTA = 2 Then
            Me.Label1.Text = Me.ListBox2.Items.Count & " Reporte(s)"
            GoTo LEE_OTRO
        End If
        If CONTROL_CARGA_LISTA = 3 Then
            Me.Label1.Text = Me.ListBox3.Items.Count & " Reporte(s)"
            GoTo LEE_OTRO
        End If
        If CONTROL_CARGA_LISTA = 4 Then
            Me.Label1.Text = Me.ListBox4.Items.Count & " Reporte(s)"
            GoTo LEE_OTRO
        End If
        If CONTROL_CARGA_LISTA = 5 Then
            Me.Label1.Text = Me.ListBox5.Items.Count & " Reporte(s)"
            GoTo LEE_OTRO
        End If
        If CONTROL_CARGA_LISTA = 6 Then
            Me.Label1.Text = Me.ListBox6.Items.Count & " Reporte(s)"
            GoTo LEE_OTRO
        End If
        If CONTROL_CARGA_LISTA = 7 Then
            Me.Label1.Text = Me.ListBox7.Items.Count & " Reporte(s)"
            GoTo LEE_OTRO
        End If
        If CONTROL_CARGA_LISTA = 8 Then
            Me.Label1.Text = Me.ListBox8.Items.Count & " Reporte(s)"
            GoTo LEE_OTRO
        End If
LEE_OTRO:
    End Sub
    Dim CONTROL_CARGA_LISTA As Integer
    Private Sub CARGAR_LISTA_COLUMNAS()
        Dim I As Integer
        Me.Cursor = Cursors.WaitCursor
        I = 0
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                Application.DoEvents()
                For I = 0 To MiDataTable.Rows.Count - 1
                    If CONTROL_CARGA_LISTA = 1 Then
                        Me.ListBox1.Items.Add(MiDataTable.Rows(I).Item(0).ToString)
                        GoTo LEE_OTRO
                    End If
                    If CONTROL_CARGA_LISTA = 2 Then
                        Me.ListBox2.Items.Add(MiDataTable.Rows(I).Item(0).ToString)
                        GoTo LEE_OTRO
                    End If
                    If CONTROL_CARGA_LISTA = 3 Then
                        Me.ListBox3.Items.Add(MiDataTable.Rows(I).Item(0).ToString)
                        GoTo LEE_OTRO
                    End If
                    If CONTROL_CARGA_LISTA = 4 Then
                        Me.ListBox4.Items.Add(MiDataTable.Rows(I).Item(0).ToString)
                        GoTo LEE_OTRO
                    End If
                    If CONTROL_CARGA_LISTA = 5 Then
                        Me.ListBox5.Items.Add(MiDataTable.Rows(I).Item(0).ToString)
                        GoTo LEE_OTRO
                    End If
                    If CONTROL_CARGA_LISTA = 6 Then
                        Me.ListBox6.Items.Add(MiDataTable.Rows(I).Item(0).ToString)
                        GoTo LEE_OTRO
                    End If
                    If CONTROL_CARGA_LISTA = 7 Then
                        Me.ListBox7.Items.Add(MiDataTable.Rows(I).Item(0).ToString)
                        GoTo LEE_OTRO
                    End If
                    If CONTROL_CARGA_LISTA = 8 Then
                        Me.ListBox8.Items.Add(MiDataTable.Rows(I).Item(0).ToString)
                        GoTo LEE_OTRO
                    End If
LEE_OTRO:
                Next
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub Frm_Seleccion_Consultas_y_Reportes_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        INFORMES = False
    End Sub
    Private Sub Frm_Seleccion_Consultas_y_Reportes_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        INFORMES = False
    End Sub
    Private Sub Frm_Seleccion_Consultas_y_Reportes_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        INFORMES = False
    End Sub
    Private Sub Frm_Seleccion_Consultas_y_Reportes_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        INFORMES = False
    End Sub
    Dim ULTIMA_LISTA_QUE_HIZO_CLIC As String
    Private Sub ListBox1_Click(sender As Object, e As EventArgs) Handles ListBox1.Click
        If Me.ListBox1.Text <> "" Then
            Call DESACTIVAR_TODO()
            CODIGO_REPORTE = Mid(Me.ListBox1.Text, 1, 5)
            If Mid(CODIGO_REPORTE, 1, 1) <> ULTIMA_LISTA_QUE_HIZO_CLIC Then
                Call LIMPIAR_LISTAS_DE_SELECCION()
                Me.Lista_Principal.Items.Clear()
            End If
            NOMBRE_REPORTE = Me.ListBox1.Text
            Label2.Text = "Orgánica\" & NOMBRE_REPORTE
            If Mid(CODIGO_REPORTE, 1, 1) = "O" Then
                '--------------- NIVELES ORGANIZACIONALES
                Me.ComboBox1.Enabled = True        'NIVEL 1
                Me.ComboBox2.Enabled = True        'NIVEL 2
                Me.CheckBox1.Enabled = True        'CHECK N2
                Me.ComboBox3.Enabled = True        'NIVEL 3
                Me.CheckBox2.Enabled = True        'CHECK N3
                Me.ComboBox4.Enabled = True        'NIVEL 4
                Me.CheckBox3.Enabled = True        'CHECK N4
                Me.ComboBox5.Enabled = True        'NIVEL 5
                Me.CheckBox4.Enabled = True        'CHECK N5
                Me.ComboBox6.Enabled = True        'NIVEL 6
                Me.CheckBox5.Enabled = True        'CHECK N6
                Me.ComboBox7.Enabled = True        'NIVEL 7
                Me.CheckBox6.Enabled = True        'CHECK N7
                Me.ComboBox8.Enabled = True        'NIVEL 8
                Me.CheckBox7.Enabled = True        'CHECK N8
                Me.Button1.Enabled = True          'BOTON NIVELES ORGANIZACIONALES
                '--------------- TIPO DE ESTRUCTURAS
                Me.RadioButton1.Enabled = True     'ESTRUCTURA MIXTA
                Me.RadioButton2.Enabled = True     'ESTRUCTURA MILITAR
                Me.RadioButton3.Enabled = True     'ESTRUCTURA PAME
                Me.RadioButton4.Enabled = True     'JEFE
                Me.Button2.Enabled = True          'BOTON TIPO DE ESTRUCTURAS
                '--------------- GRADOS, CARGOS, OTROS
                Me.ComboBox9.Enabled = True        'GRADO REAL
                Me.Button3.Enabled = True          'BOTON GR
                Me.ComboBox10.Enabled = True       'GRADO PLANTILLA
                Me.Button4.Enabled = True          'BOTON GP
                Me.ComboBox11.Enabled = True       'SITUACION DEL CARGO
                Me.Button5.Enabled = True          'BOTON SITUACION CARGO
                Me.ComboBox12.Enabled = True       'CARGO
                Me.Button6.Enabled = True          'BOTON CARGO
                Me.ComboBox13.Enabled = True       'ESTABLECIMIENTO
                Me.Button7.Enabled = True          'BOTON ESTABLECIMIENTO
                '--------------- FECHAS
                Me.RadioButton5.Enabled = True     'FECHA DE INGRESO PAME
                Me.RadioButton5.Checked = True
                Me.RadioButton6.Enabled = True     'FECHA DE INGRESO EN
                Me.RadioButton7.Enabled = True     'FECHA DE NACIMIENTO
                Me.DateTimePicker1.Enabled = True  'DESDE
                Me.DateTimePicker2.Enabled = True  'HASTA
                Me.Button11.Enabled = True         'BOTON FECHAS
                '--------------- OTROS FILTROS
                'Me.ComboBox19.Enabled = True       'FUNCIONALIDAD DE EMPLEADO
                'Me.Button14.Enabled = True         'BOTON FUNCIONALIDAD DE EMPLEADO
                'Me.ComboBox20.Enabled = True       'CAUSA REAL DE EVENTUALIDAD
                'Me.Button15.Enabled = True         'BOTON CAUSA REAL DE EVENTUALIDAD
                'Me.ComboBox21.Enabled = True       'TIPO DE EVENTUALIDAD
                'Me.Button16.Enabled = True         'BOTON TIPO DE EVENTUALIDAD
                'Me.ComboBox22.Enabled = True       'ESTUDIOS
                'Me.Button17.Enabled = True         'BOTON ESTUDIOS
                'Me.ComboBox23.Enabled = True       'NIVEL DE COMPETENCIA
                'Me.Button18.Enabled = True         'BOTON NIVEL DE COMPETENCIA
                'Me.ComboBox26.Enabled = True       'CLASIFICACION DEL EMPLEADO
                'Me.Button21.Enabled = True         'BOTON CLASIFICACION
                'Me.TextBox1.Enabled = True         'CODIGO DEL EMPLEADO
                'Me.Button22.Enabled = True         'BOTON CODIGO EMPLEADO
                'Me.RadioButton12.Enabled = True    'EDAD
                'Me.RadioButton13.Enabled = True    'ANTIGUEDAD PAME
                'Me.RadioButton14.Enabled = True    'ANTIGUEDAD EN
                'Me.TextBox2.Enabled = True         'ANTIGUEDADES A
                'Me.TextBox3.Enabled = True         'ANTIGUEDADES B
                'Me.Button23.Enabled = True         'BOTON ANTIGUEDADES
                'Me.TextBox4.Enabled = True         'CONTRATOS
                'Me.Button24.Enabled = True         'BOTON CONTRATOS
                '--------------- CONDECORACIONES
                'Me.ComboBox27.Enabled = True      'CLASIFICACION
                'Me.Button25.Enabled = True        'BOTON CLASIFICACION
                'Me.ComboBox28.Enabled = True      'TIPO
                'Me.Button26.Enabled = True        'BOTON TIPO
                '--------------- SITUACIONES
                'Me.ComboBox29.Enabled = True      'TIPO SITUACION
                'Me.Button27.Enabled = True        'BOTON TIPO SITUACION
                'Me.ComboBox30.Enabled = True      'SITUACION
                'Me.Button28.Enabled = True        'BOTON SITUACION
                'Me.DateTimePicker3.Enabled = True 'DESDE
                'Me.DateTimePicker4.Enabled = True 'HASTA
                'Me.Button29.Enabled = True        'BOTON FECHAS

                ULTIMA_LISTA_QUE_HIZO_CLIC = Mid(CODIGO_REPORTE, 1, 1)
                GoTo SALIR
            End If
SALIR:
        End If
    End Sub
    Private Sub ListBox2_Click(sender As Object, e As EventArgs) Handles ListBox2.Click
        If Me.ListBox2.Text <> "" Then
            Call DESACTIVAR_TODO()
            CODIGO_REPORTE = Mid(Me.ListBox2.Text, 1, 5)
            If Mid(CODIGO_REPORTE, 1, 1) <> ULTIMA_LISTA_QUE_HIZO_CLIC Then
                Call LIMPIAR_LISTAS_DE_SELECCION()
                Me.Lista_Principal.Items.Clear()
            End If
            NOMBRE_REPORTE = Me.ListBox2.Text
            Label2.Text = NOMBRE_REPORTE
            Label2.Text = "Movimientos\" & NOMBRE_REPORTE
            If Mid(CODIGO_REPORTE, 1, 1) = "M" Then
                '--------------- NIVELES ORGANIZACIONALES
                Me.ComboBox1.Enabled = True        'NIVEL 1
                Me.ComboBox2.Enabled = True        'NIVEL 2
                Me.CheckBox1.Enabled = True        'CHECK N2
                Me.ComboBox3.Enabled = True        'NIVEL 3
                Me.CheckBox2.Enabled = True        'CHECK N3
                Me.ComboBox4.Enabled = True        'NIVEL 4
                Me.CheckBox3.Enabled = True        'CHECK N4
                Me.ComboBox5.Enabled = True        'NIVEL 5
                Me.CheckBox4.Enabled = True        'CHECK N5
                Me.ComboBox6.Enabled = True        'NIVEL 6
                Me.CheckBox5.Enabled = True        'CHECK N6
                Me.ComboBox7.Enabled = True        'NIVEL 7
                Me.CheckBox6.Enabled = True        'CHECK N7
                Me.ComboBox8.Enabled = True        'NIVEL 8
                Me.CheckBox7.Enabled = True        'CHECK N8
                Me.Button1.Enabled = True          'BOTON NIVELES ORGANIZACIONALES
                '--------------- TIPO DE ESTRUCTURAS
                Me.RadioButton1.Enabled = True     'ESTRUCTURA MIXTA
                Me.RadioButton2.Enabled = True     'ESTRUCTURA MILITAR
                Me.RadioButton3.Enabled = True     'ESTRUCTURA PAME
                Me.RadioButton4.Enabled = True     'JEFE
                Me.Button2.Enabled = True          'BOTON TIPO DE ESTRUCTURAS
                '--------------- GRADOS, CARGOS, OTROS
                Me.ComboBox9.Enabled = True        'GRADO REAL
                Me.Button3.Enabled = True          'BOTON GR
                Me.ComboBox10.Enabled = True       'GRADO PLANTILLA
                Me.Button4.Enabled = True          'BOTON GP
                'Me.ComboBox11.Enabled = True       'SITUACION DEL CARGO
                'Me.Button5.Enabled = True          'BOTON SITUACION CARGO
                Me.ComboBox12.Enabled = True       'CARGO
                Me.Button6.Enabled = True          'BOTON CARGO
                'Me.ComboBox13.Enabled = True       'ESTABLECIMIENTO
                'Me.Button7.Enabled = True          'BOTON ESTABLECIMIENTO
                '--------------- MOVIMIENTOS
                Me.ComboBox14.Enabled = True       'TIPO DE MOVIMIENTO
                Me.Button8.Enabled = True          'BOTON TIPO MOV
                Me.ComboBox15.Enabled = True       'CAUSAL REAL 
                Me.Button9.Enabled = True          'BOTON CAUSAL REAL
                Me.ComboBox16.Enabled = True       'CAUSAL LEGAL
                Me.Button10.Enabled = True         'BOTON CAUSAL LEGAL
                '--------------- FECHAS
                Me.RadioButton11.Enabled = True    'FECHA DE MOVIMIENTOS
                Me.RadioButton11.Checked = True
                Me.DateTimePicker1.Enabled = True  'DESDE
                Me.DateTimePicker2.Enabled = True  'HASTA
                Me.Button11.Enabled = True         'BOTON FECHAS
                '--------------- OTROS FILTROS
                'Me.ComboBox19.Enabled = True       'FUNCIONALIDAD DE EMPLEADO
                'Me.Button14.Enabled = True         'BOTON FUNCIONALIDAD DE EMPLEADO
                'Me.ComboBox20.Enabled = True       'CAUSA REAL DE EVENTUALIDAD
                'Me.Button15.Enabled = True         'BOTON CAUSA REAL DE EVENTUALIDAD
                'Me.ComboBox21.Enabled = True       'TIPO DE EVENTUALIDAD
                'Me.Button16.Enabled = True         'BOTON TIPO DE EVENTUALIDAD
                'Me.ComboBox22.Enabled = True       'ESTUDIOS
                'Me.Button17.Enabled = True         'BOTON ESTUDIOS
                'Me.ComboBox23.Enabled = True       'NIVEL DE COMPETENCIA
                'Me.Button18.Enabled = True         'BOTON NIVEL DE COMPETENCIA
                'Me.ComboBox26.Enabled = True       'CLASIFICACION DEL EMPLEADO
                'Me.Button21.Enabled = True         'BOTON CLASIFICACION
                'Me.TextBox1.Enabled = True         'CODIGO DEL EMPLEADO
                'Me.Button22.Enabled = True         'BOTON CODIGO EMPLEADO
                'Me.RadioButton12.Enabled = True    'EDAD
                'Me.RadioButton13.Enabled = True    'ANTIGUEDAD PAME
                'Me.RadioButton14.Enabled = True    'ANTIGUEDAD EN
                'Me.TextBox2.Enabled = True         'ANTIGUEDADES A
                'Me.TextBox3.Enabled = True         'ANTIGUEDADES B
                'Me.Button23.Enabled = True         'BOTON ANTIGUEDADES
                'Me.TextBox4.Enabled = True         'CONTRATOS
                'Me.Button24.Enabled = True         'BOTON CONTRATOS
                '--------------- CONDECORACIONES
                'Me.ComboBox27.Enabled = True      'CLASIFICACION
                'Me.Button25.Enabled = True        'BOTON CLASIFICACION
                'Me.ComboBox28.Enabled = True      'TIPO
                'Me.Button26.Enabled = True        'BOTON TIPO
                '--------------- SITUACIONES
                'Me.ComboBox29.Enabled = True      'TIPO SITUACION
                'Me.Button27.Enabled = True        'BOTON TIPO SITUACION
                'Me.ComboBox30.Enabled = True      'SITUACION
                'Me.Button28.Enabled = True        'BOTON SITUACION
                'Me.DateTimePicker3.Enabled = True 'DESDE
                'Me.DateTimePicker4.Enabled = True 'HASTA
                'Me.Button29.Enabled = True        'BOTON FECHAS

                ULTIMA_LISTA_QUE_HIZO_CLIC = Mid(CODIGO_REPORTE, 1, 1)
                GoTo SALIR
            End If
SALIR:

        End If
    End Sub
    Private Sub ListBox3_Click(sender As Object, e As EventArgs) Handles ListBox3.Click
        If Me.ListBox3.Text <> "" Then
            Call DESACTIVAR_TODO()
            CODIGO_REPORTE = Mid(Me.ListBox3.Text, 1, 5)
            If Mid(CODIGO_REPORTE, 1, 1) <> ULTIMA_LISTA_QUE_HIZO_CLIC Then
                Call LIMPIAR_LISTAS_DE_SELECCION()
                Me.Lista_Principal.Items.Clear()
            End If
            NOMBRE_REPORTE = Me.ListBox3.Text
            Label2.Text = NOMBRE_REPORTE
            Label2.Text = "Aspirantes (RSC)\" & NOMBRE_REPORTE
            If Mid(CODIGO_REPORTE, 1, 1) = "A" Then
                '--------------- FECHAS
                Me.RadioButton10.Enabled = True    'FECHA DE PROCESO DE ASPIRANTE
                Me.RadioButton10.Checked = True
                Me.DateTimePicker1.Enabled = True  'DESDE
                Me.DateTimePicker2.Enabled = True  'HASTA
                Me.Button11.Enabled = True         'BOTON FECHAS
                '--------------- ASPIRANTES
                Me.ComboBox17.Enabled = True       'PROCESO
                Me.Button12.Enabled = True         'BOTON  PROCESO
                Me.ComboBox18.Enabled = True       'RESULTADO
                Me.Button13.Enabled = True         'BOTON RESULTADO

                ULTIMA_LISTA_QUE_HIZO_CLIC = Mid(CODIGO_REPORTE, 1, 1)
                GoTo SALIR
            End If
SALIR:
        End If
    End Sub
    Private Sub ListBox4_Click(sender As Object, e As EventArgs) Handles ListBox4.Click
        If Me.ListBox4.Text <> "" Then
            Call DESACTIVAR_TODO()
            CODIGO_REPORTE = Mid(Me.ListBox4.Text, 1, 5)
            If Mid(CODIGO_REPORTE, 1, 1) <> ULTIMA_LISTA_QUE_HIZO_CLIC Then
                Call LIMPIAR_LISTAS_DE_SELECCION()
                Me.Lista_Principal.Items.Clear()
            End If
            NOMBRE_REPORTE = Me.ListBox4.Text
            Label2.Text = NOMBRE_REPORTE
            Label2.Text = "Expedientes\" & NOMBRE_REPORTE
            If (CODIGO_REPORTE = "E0001") Or (CODIGO_REPORTE = "E0015") Then                      'E0001 - Expediente de Empleado (ACTIVOS - INACTIVOS)
                '--------------- NIVELES ORGANIZACIONALES
                Me.ComboBox1.Enabled = True        'NIVEL 1
                Me.ComboBox2.Enabled = True        'NIVEL 2
                Me.CheckBox1.Enabled = True        'CHECK N2
                Me.ComboBox3.Enabled = True        'NIVEL 3
                Me.CheckBox2.Enabled = True        'CHECK N3
                Me.ComboBox4.Enabled = True        'NIVEL 4
                Me.CheckBox3.Enabled = True        'CHECK N4
                Me.ComboBox5.Enabled = True        'NIVEL 5
                Me.CheckBox4.Enabled = True        'CHECK N5
                Me.ComboBox6.Enabled = True        'NIVEL 6
                Me.CheckBox5.Enabled = True        'CHECK N6
                Me.ComboBox7.Enabled = True        'NIVEL 7
                Me.CheckBox6.Enabled = True        'CHECK N7
                Me.ComboBox8.Enabled = True        'NIVEL 8
                Me.CheckBox7.Enabled = True        'CHECK N8
                Me.Button1.Enabled = True          'BOTON NIVELES ORGANIZACIONALES
                '--------------- TIPO DE ESTRUCTURAS
                Me.RadioButton1.Enabled = True     'ESTRUCTURA MIXTA
                Me.RadioButton2.Enabled = True     'ESTRUCTURA MILITAR
                Me.RadioButton3.Enabled = True     'ESTRUCTURA PAME
                Me.RadioButton4.Enabled = True     'JEFE
                Me.Button2.Enabled = True          'BOTON TIPO DE ESTRUCTURAS
                '--------------- GRADOS, CARGOS, OTROS
                Me.ComboBox9.Enabled = True        'GRADO REAL
                Me.Button3.Enabled = True          'BOTON GR
                Me.ComboBox10.Enabled = True       'GRADO PLANTILLA
                Me.Button4.Enabled = True          'BOTON GP
                Me.ComboBox11.Enabled = True       'SITUACION DEL CARGO
                Me.Button5.Enabled = True          'BOTON SITUACION CARGO
                Me.ComboBox12.Enabled = True       'CARGO
                Me.Button6.Enabled = True          'BOTON CARGO
                Me.ComboBox13.Enabled = True       'ESTABLECIMIENTO
                Me.Button7.Enabled = True          'BOTON ESTABLECIMIENTO
                '--------------- FECHAS
                Me.RadioButton5.Enabled = True     'FECHA DE INGRESO PAME
                Me.RadioButton5.Checked = True
                Me.RadioButton6.Enabled = True     'FECHA DE INGRESO EN
                Me.RadioButton7.Enabled = True     'FECHA DE NACIMIENTO
                Me.DateTimePicker1.Enabled = True  'DESDE
                Me.DateTimePicker2.Enabled = True  'HASTA
                Me.Button11.Enabled = True         'BOTON FECHAS
                '--------------- OTROS FILTROS
                Me.ComboBox19.Enabled = True       'FUNCIONALIDAD DE EMPLEADO
                Me.Button14.Enabled = True         'BOTON FUNCIONALIDAD DE EMPLEADO
                'Me.ComboBox20.Enabled = True       'CAUSA REAL DE EVENTUALIDAD
                'Me.Button15.Enabled = True         'BOTON CAUSA REAL DE EVENTUALIDAD
                'Me.ComboBox21.Enabled = True       'TIPO DE EVENTUALIDAD
                'Me.Button16.Enabled = True         'BOTON TIPO DE EVENTUALIDAD
                'Me.ComboBox22.Enabled = True       'ESTUDIOS
                'Me.Button17.Enabled = True         'BOTON ESTUDIOS
                Me.ComboBox23.Enabled = True       'NIVEL DE COMPETENCIA
                Me.Button18.Enabled = True         'BOTON NIVEL DE COMPETENCIA
                'Me.ComboBox26.Enabled = True       'CLASIFICACION DEL EMPLEADO
                'Me.Button21.Enabled = True         'BOTON CLASIFICACION
                Me.TextBox1.Enabled = True         'CODIGO DEL EMPLEADO
                Me.Button22.Enabled = True         'BOTON CODIGO EMPLEADO
                Me.RadioButton12.Enabled = True    'EDAD
                'Me.RadioButton13.Enabled = True    'ANTIGUEDAD PAME
                'Me.RadioButton14.Enabled = True    'ANTIGUEDAD EN
                'Me.TextBox2.Enabled = True         'ANTIGUEDADES A
                'Me.TextBox3.Enabled = True         'ANTIGUEDADES B
                'Me.Button23.Enabled = True         'BOTON ANTIGUEDADES
                'Me.TextBox4.Enabled = True         'CONTRATOS
                'Me.Button24.Enabled = True         'BOTON CONTRATOS
                '--------------- CONDECORACIONES
                'Me.ComboBox27.Enabled = True      'CLASIFICACION
                'Me.Button25.Enabled = True        'BOTON CLASIFICACION
                'Me.ComboBox28.Enabled = True      'TIPO
                'Me.Button26.Enabled = True        'BOTON TIPO
                '--------------- SITUACIONES
                'Me.ComboBox29.Enabled = True      'TIPO SITUACION
                'Me.Button27.Enabled = True        'BOTON TIPO SITUACION
                'Me.ComboBox30.Enabled = True      'SITUACION
                'Me.Button28.Enabled = True        'BOTON SITUACION
                'Me.DateTimePicker3.Enabled = True 'DESDE
                'Me.DateTimePicker4.Enabled = True 'HASTA
                'Me.Button29.Enabled = True        'BOTON FECHAS

                ULTIMA_LISTA_QUE_HIZO_CLIC = Mid(CODIGO_REPORTE, 1, 1)
                GoTo SALIR
            End If
            If (CODIGO_REPORTE = "E0002") Then                      'E0002 - Empleados por Antiguedad (Edad, Antiguedad PAME, Antiguedad EN)
                '--------------- NIVELES ORGANIZACIONALES
                Me.ComboBox1.Enabled = True        'NIVEL 1
                Me.ComboBox2.Enabled = True        'NIVEL 2
                Me.CheckBox1.Enabled = True        'CHECK N2
                Me.ComboBox3.Enabled = True        'NIVEL 3
                Me.CheckBox2.Enabled = True        'CHECK N3
                Me.ComboBox4.Enabled = True        'NIVEL 4
                Me.CheckBox3.Enabled = True        'CHECK N4
                Me.ComboBox5.Enabled = True        'NIVEL 5
                Me.CheckBox4.Enabled = True        'CHECK N5
                Me.ComboBox6.Enabled = True        'NIVEL 6
                Me.CheckBox5.Enabled = True        'CHECK N6
                Me.ComboBox7.Enabled = True        'NIVEL 7
                Me.CheckBox6.Enabled = True        'CHECK N7
                Me.ComboBox8.Enabled = True        'NIVEL 8
                Me.CheckBox7.Enabled = True        'CHECK N8
                Me.Button1.Enabled = True          'BOTON NIVELES ORGANIZACIONALES
                '--------------- TIPO DE ESTRUCTURAS
                Me.RadioButton1.Enabled = True     'ESTRUCTURA MIXTA
                Me.RadioButton2.Enabled = True     'ESTRUCTURA MILITAR
                Me.RadioButton3.Enabled = True     'ESTRUCTURA PAME
                Me.RadioButton4.Enabled = True     'JEFE
                Me.Button2.Enabled = True          'BOTON TIPO DE ESTRUCTURAS
                '--------------- GRADOS, CARGOS, OTROS
                Me.ComboBox9.Enabled = True        'GRADO REAL
                Me.Button3.Enabled = True          'BOTON GR
                Me.ComboBox10.Enabled = True       'GRADO PLANTILLA
                Me.Button4.Enabled = True          'BOTON GP
                Me.ComboBox11.Enabled = True       'SITUACION DEL CARGO
                Me.Button5.Enabled = True          'BOTON SITUACION CARGO
                Me.ComboBox12.Enabled = True       'CARGO
                Me.Button6.Enabled = True          'BOTON CARGO
                Me.ComboBox13.Enabled = True       'ESTABLECIMIENTO
                Me.Button7.Enabled = True          'BOTON ESTABLECIMIENTO
                '--------------- FECHAS
                Me.RadioButton5.Enabled = True     'FECHA DE INGRESO PAME
                Me.RadioButton5.Checked = True
                Me.RadioButton6.Enabled = True     'FECHA DE INGRESO EN
                Me.RadioButton7.Enabled = True     'FECHA DE NACIMIENTO
                Me.DateTimePicker1.Enabled = True  'DESDE
                Me.DateTimePicker2.Enabled = True  'HASTA
                Me.Button11.Enabled = True         'BOTON FECHAS
                '--------------- OTROS FILTROS
                'Me.ComboBox19.Enabled = True       'FUNCIONALIDAD DE EMPLEADO
                'Me.Button14.Enabled = True         'BOTON FUNCIONALIDAD DE EMPLEADO
                'Me.ComboBox20.Enabled = True       'CAUSA REAL DE EVENTUALIDAD
                'Me.Button15.Enabled = True         'BOTON CAUSA REAL DE EVENTUALIDAD
                'Me.ComboBox21.Enabled = True       'TIPO DE EVENTUALIDAD
                'Me.Button16.Enabled = True         'BOTON TIPO DE EVENTUALIDAD
                'Me.ComboBox22.Enabled = True       'ESTUDIOS
                'Me.Button17.Enabled = True         'BOTON ESTUDIOS
                'Me.ComboBox23.Enabled = True       'NIVEL DE COMPETENCIA
                'Me.Button18.Enabled = True         'BOTON NIVEL DE COMPETENCIA
                'Me.ComboBox26.Enabled = True       'CLASIFICACION DEL EMPLEADO
                'Me.Button21.Enabled = True         'BOTON CLASIFICACION
                Me.TextBox1.Enabled = True         'CODIGO DEL EMPLEADO
                Me.Button22.Enabled = True         'BOTON CODIGO EMPLEADO
                Me.RadioButton12.Enabled = True    'EDAD
                Me.RadioButton13.Enabled = True    'ANTIGUEDAD PAME
                Me.RadioButton14.Enabled = True    'ANTIGUEDAD EN
                Me.TextBox2.Enabled = True         'ANTIGUEDADES A
                Me.TextBox3.Enabled = True         'ANTIGUEDADES B
                Me.Button23.Enabled = True         'BOTON ANTIGUEDADES
                'Me.TextBox4.Enabled = True         'CONTRATOS
                'Me.Button24.Enabled = True         'BOTON CONTRATOS
                '--------------- CONDECORACIONES
                'Me.ComboBox27.Enabled = True      'CLASIFICACION
                'Me.Button25.Enabled = True        'BOTON CLASIFICACION
                'Me.ComboBox28.Enabled = True      'TIPO
                'Me.Button26.Enabled = True        'BOTON TIPO
                '--------------- SITUACIONES
                'Me.ComboBox29.Enabled = True      'TIPO SITUACION
                'Me.Button27.Enabled = True        'BOTON TIPO SITUACION
                'Me.ComboBox30.Enabled = True      'SITUACION
                'Me.Button28.Enabled = True        'BOTON SITUACION
                'Me.DateTimePicker3.Enabled = True 'DESDE
                'Me.DateTimePicker4.Enabled = True 'HASTA
                'Me.Button29.Enabled = True        'BOTON FECHAS

                ULTIMA_LISTA_QUE_HIZO_CLIC = Mid(CODIGO_REPORTE, 1, 1)
                GoTo SALIR
            End If
            If (CODIGO_REPORTE = "E0003") Then                      'E0003 - Lista de Empleados por Clasificacion
                '--------------- NIVELES ORGANIZACIONALES
                Me.ComboBox1.Enabled = True        'NIVEL 1
                Me.ComboBox2.Enabled = True        'NIVEL 2
                Me.CheckBox1.Enabled = True        'CHECK N2
                Me.ComboBox3.Enabled = True        'NIVEL 3
                Me.CheckBox2.Enabled = True        'CHECK N3
                Me.ComboBox4.Enabled = True        'NIVEL 4
                Me.CheckBox3.Enabled = True        'CHECK N4
                Me.ComboBox5.Enabled = True        'NIVEL 5
                Me.CheckBox4.Enabled = True        'CHECK N5
                Me.ComboBox6.Enabled = True        'NIVEL 6
                Me.CheckBox5.Enabled = True        'CHECK N6
                Me.ComboBox7.Enabled = True        'NIVEL 7
                Me.CheckBox6.Enabled = True        'CHECK N7
                Me.ComboBox8.Enabled = True        'NIVEL 8
                Me.CheckBox7.Enabled = True        'CHECK N8
                Me.Button1.Enabled = True          'BOTON NIVELES ORGANIZACIONALES
                '--------------- TIPO DE ESTRUCTURAS
                Me.RadioButton1.Enabled = True     'ESTRUCTURA MIXTA
                Me.RadioButton2.Enabled = True     'ESTRUCTURA MILITAR
                Me.RadioButton3.Enabled = True     'ESTRUCTURA PAME
                Me.RadioButton4.Enabled = True     'JEFE
                Me.Button2.Enabled = True          'BOTON TIPO DE ESTRUCTURAS
                '--------------- GRADOS, CARGOS, OTROS
                Me.ComboBox9.Enabled = True        'GRADO REAL
                Me.Button3.Enabled = True          'BOTON GR
                Me.ComboBox10.Enabled = True       'GRADO PLANTILLA
                Me.Button4.Enabled = True          'BOTON GP
                Me.ComboBox11.Enabled = True       'SITUACION DEL CARGO
                Me.Button5.Enabled = True          'BOTON SITUACION CARGO
                Me.ComboBox12.Enabled = True       'CARGO
                Me.Button6.Enabled = True          'BOTON CARGO
                Me.ComboBox13.Enabled = True       'ESTABLECIMIENTO
                Me.Button7.Enabled = True          'BOTON ESTABLECIMIENTO
                '--------------- FECHAS
                Me.RadioButton5.Enabled = True     'FECHA DE INGRESO PAME
                Me.RadioButton5.Checked = True
                Me.RadioButton6.Enabled = True     'FECHA DE INGRESO EN
                Me.RadioButton7.Enabled = True     'FECHA DE NACIMIENTO
                Me.DateTimePicker1.Enabled = True  'DESDE
                Me.DateTimePicker2.Enabled = True  'HASTA
                Me.Button11.Enabled = True         'BOTON FECHAS
                '--------------- OTROS FILTROS
                Me.ComboBox19.Enabled = True       'FUNCIONALIDAD DE EMPLEADO
                Me.Button14.Enabled = True         'BOTON FUNCIONALIDAD DE EMPLEADO
                'Me.ComboBox20.Enabled = True       'CAUSA REAL DE EVENTUALIDAD
                'Me.Button15.Enabled = True         'BOTON CAUSA REAL DE EVENTUALIDAD
                'Me.ComboBox21.Enabled = True       'TIPO DE EVENTUALIDAD
                'Me.Button16.Enabled = True         'BOTON TIPO DE EVENTUALIDAD
                'Me.ComboBox22.Enabled = True       'ESTUDIOS
                'Me.Button17.Enabled = True         'BOTON ESTUDIOS
                'Me.ComboBox23.Enabled = True       'NIVEL DE COMPETENCIA
                'Me.Button18.Enabled = True         'BOTON NIVEL DE COMPETENCIA
                Me.ComboBox26.Enabled = True       'CLASIFICACION DEL EMPLEADO
                Me.Button21.Enabled = True         'BOTON CLASIFICACION
                Me.TextBox1.Enabled = True         'CODIGO DEL EMPLEADO
                Me.Button22.Enabled = True         'BOTON CODIGO EMPLEADO
                'Me.RadioButton12.Enabled = True    'EDAD
                'Me.RadioButton13.Enabled = True    'ANTIGUEDAD PAME
                'Me.RadioButton14.Enabled = True    'ANTIGUEDAD EN
                'Me.TextBox2.Enabled = True         'ANTIGUEDADES A
                'Me.TextBox3.Enabled = True         'ANTIGUEDADES B
                'Me.Button23.Enabled = True         'BOTON ANTIGUEDADES
                'Me.TextBox4.Enabled = True         'CONTRATOS
                'Me.Button24.Enabled = True         'BOTON CONTRATOS
                '--------------- CONDECORACIONES
                'Me.ComboBox27.Enabled = True      'CLASIFICACION
                'Me.Button25.Enabled = True        'BOTON CLASIFICACION
                'Me.ComboBox28.Enabled = True      'TIPO
                'Me.Button26.Enabled = True        'BOTON TIPO
                '--------------- SITUACIONES
                'Me.ComboBox29.Enabled = True      'TIPO SITUACION
                'Me.Button27.Enabled = True        'BOTON TIPO SITUACION
                'Me.ComboBox30.Enabled = True      'SITUACION
                'Me.Button28.Enabled = True        'BOTON SITUACION
                'Me.DateTimePicker3.Enabled = True 'DESDE
                'Me.DateTimePicker4.Enabled = True 'HASTA
                'Me.Button29.Enabled = True        'BOTON FECHAS

                ULTIMA_LISTA_QUE_HIZO_CLIC = Mid(CODIGO_REPORTE, 1, 1)
                GoTo SALIR
            End If
            If (CODIGO_REPORTE = "E0004") Then                      'E0004 - Historico de Condecoraciones por Empleado
                '--------------- NIVELES ORGANIZACIONALES
                Me.ComboBox1.Enabled = True        'NIVEL 1
                Me.ComboBox2.Enabled = True        'NIVEL 2
                Me.CheckBox1.Enabled = True        'CHECK N2
                Me.ComboBox3.Enabled = True        'NIVEL 3
                Me.CheckBox2.Enabled = True        'CHECK N3
                Me.ComboBox4.Enabled = True        'NIVEL 4
                Me.CheckBox3.Enabled = True        'CHECK N4
                Me.ComboBox5.Enabled = True        'NIVEL 5
                Me.CheckBox4.Enabled = True        'CHECK N5
                Me.ComboBox6.Enabled = True        'NIVEL 6
                Me.CheckBox5.Enabled = True        'CHECK N6
                Me.ComboBox7.Enabled = True        'NIVEL 7
                Me.CheckBox6.Enabled = True        'CHECK N7
                Me.ComboBox8.Enabled = True        'NIVEL 8
                Me.CheckBox7.Enabled = True        'CHECK N8
                Me.Button1.Enabled = True          'BOTON NIVELES ORGANIZACIONALES
                '--------------- TIPO DE ESTRUCTURAS
                Me.RadioButton1.Enabled = True     'ESTRUCTURA MIXTA
                Me.RadioButton2.Enabled = True     'ESTRUCTURA MILITAR
                Me.RadioButton3.Enabled = True     'ESTRUCTURA PAME
                Me.RadioButton4.Enabled = True     'JEFE
                Me.Button2.Enabled = True          'BOTON TIPO DE ESTRUCTURAS
                '--------------- GRADOS, CARGOS, OTROS
                Me.ComboBox9.Enabled = True        'GRADO REAL
                Me.Button3.Enabled = True          'BOTON GR
                Me.ComboBox10.Enabled = True       'GRADO PLANTILLA
                Me.Button4.Enabled = True          'BOTON GP
                Me.ComboBox11.Enabled = True       'SITUACION DEL CARGO
                Me.Button5.Enabled = True          'BOTON SITUACION CARGO
                Me.ComboBox12.Enabled = True       'CARGO
                Me.Button6.Enabled = True          'BOTON CARGO
                Me.ComboBox13.Enabled = True       'ESTABLECIMIENTO
                Me.Button7.Enabled = True          'BOTON ESTABLECIMIENTO
                '--------------- FECHAS
                Me.RadioButton5.Enabled = True     'FECHA DE INGRESO PAME
                Me.RadioButton5.Checked = True
                Me.RadioButton6.Enabled = True     'FECHA DE INGRESO EN
                Me.RadioButton7.Enabled = True     'FECHA DE NACIMIENTO
                Me.DateTimePicker1.Enabled = True  'DESDE
                Me.DateTimePicker2.Enabled = True  'HASTA
                Me.Button11.Enabled = True         'BOTON FECHAS
                '--------------- OTROS FILTROS
                Me.ComboBox19.Enabled = True       'FUNCIONALIDAD DE EMPLEADO
                Me.Button14.Enabled = True         'BOTON FUNCIONALIDAD DE EMPLEADO
                'Me.ComboBox20.Enabled = True       'CAUSA REAL DE EVENTUALIDAD
                'Me.Button15.Enabled = True         'BOTON CAUSA REAL DE EVENTUALIDAD
                'Me.ComboBox21.Enabled = True       'TIPO DE EVENTUALIDAD
                'Me.Button16.Enabled = True         'BOTON TIPO DE EVENTUALIDAD
                'Me.ComboBox22.Enabled = True       'ESTUDIOS
                'Me.Button17.Enabled = True         'BOTON ESTUDIOS
                Me.ComboBox23.Enabled = True       'NIVEL DE COMPETENCIA
                Me.Button18.Enabled = True         'BOTON NIVEL DE COMPETENCIA
                'Me.ComboBox26.Enabled = True       'CLASIFICACION DEL EMPLEADO
                'Me.Button21.Enabled = True         'BOTON CLASIFICACION
                Me.TextBox1.Enabled = True         'CODIGO DEL EMPLEADO
                Me.Button22.Enabled = True         'BOTON CODIGO EMPLEADO
                'Me.RadioButton12.Enabled = True    'EDAD
                'Me.RadioButton13.Enabled = True    'ANTIGUEDAD PAME
                'Me.RadioButton14.Enabled = True    'ANTIGUEDAD EN
                'Me.TextBox2.Enabled = True         'ANTIGUEDADES A
                'Me.TextBox3.Enabled = True         'ANTIGUEDADES B
                'Me.Button23.Enabled = True         'BOTON ANTIGUEDADES
                'Me.TextBox4.Enabled = True         'CONTRATOS
                'Me.Button24.Enabled = True         'BOTON CONTRATOS
                '--------------- CONDECORACIONES
                Me.ComboBox27.Enabled = True        'CLASIFICACION
                Me.Button25.Enabled = True          'BOTON CLASIFICACION
                Me.ComboBox28.Enabled = True        'TIPO
                Me.Button26.Enabled = True          'BOTON TIPO
                '--------------- SITUACIONES
                'Me.ComboBox29.Enabled = True      'TIPO SITUACION
                'Me.Button27.Enabled = True        'BOTON TIPO SITUACION
                'Me.ComboBox30.Enabled = True      'SITUACION
                'Me.Button28.Enabled = True        'BOTON SITUACION
                'Me.DateTimePicker3.Enabled = True 'DESDE
                'Me.DateTimePicker4.Enabled = True 'HASTA
                'Me.Button29.Enabled = True        'BOTON FECHAS

                ULTIMA_LISTA_QUE_HIZO_CLIC = Mid(CODIGO_REPORTE, 1, 1)
                GoTo SALIR
            End If
            If (CODIGO_REPORTE = "E0005") Then                      'E0005 - Curva y Talla de Empleados por Niveles Organizacionales
                '--------------- NIVELES ORGANIZACIONALES
                Me.ComboBox1.Enabled = True        'NIVEL 1
                Me.ComboBox2.Enabled = True        'NIVEL 2
                Me.CheckBox1.Enabled = True        'CHECK N2
                Me.ComboBox3.Enabled = True        'NIVEL 3
                Me.CheckBox2.Enabled = True        'CHECK N3
                Me.ComboBox4.Enabled = True        'NIVEL 4
                Me.CheckBox3.Enabled = True        'CHECK N4
                Me.ComboBox5.Enabled = True        'NIVEL 5
                Me.CheckBox4.Enabled = True        'CHECK N5
                Me.ComboBox6.Enabled = True        'NIVEL 6
                Me.CheckBox5.Enabled = True        'CHECK N6
                Me.ComboBox7.Enabled = True        'NIVEL 7
                Me.CheckBox6.Enabled = True        'CHECK N7
                Me.ComboBox8.Enabled = True        'NIVEL 8
                Me.CheckBox7.Enabled = True        'CHECK N8
                Me.Button1.Enabled = True          'BOTON NIVELES ORGANIZACIONALES
                '--------------- TIPO DE ESTRUCTURAS
                Me.RadioButton1.Enabled = True     'ESTRUCTURA MIXTA
                Me.RadioButton2.Enabled = True     'ESTRUCTURA MILITAR
                Me.RadioButton3.Enabled = True     'ESTRUCTURA PAME
                Me.RadioButton4.Enabled = True     'JEFE
                Me.Button2.Enabled = True          'BOTON TIPO DE ESTRUCTURAS
                '--------------- GRADOS, CARGOS, OTROS
                Me.ComboBox9.Enabled = True        'GRADO REAL
                Me.Button3.Enabled = True          'BOTON GR
                'Me.ComboBox10.Enabled = True       'GRADO PLANTILLA
                'Me.Button4.Enabled = True          'BOTON GP
                Me.ComboBox11.Enabled = True       'SITUACION DEL CARGO
                Me.Button5.Enabled = True          'BOTON SITUACION CARGO
                Me.ComboBox12.Enabled = True       'CARGO
                Me.Button6.Enabled = True          'BOTON CARGO
                Me.ComboBox13.Enabled = True       'ESTABLECIMIENTO
                Me.Button7.Enabled = True          'BOTON ESTABLECIMIENTO
                '--------------- FECHAS
                Me.RadioButton5.Enabled = True     'FECHA DE INGRESO PAME
                Me.RadioButton5.Checked = True
                'Me.RadioButton6.Enabled = True     'FECHA DE INGRESO EN
                'Me.RadioButton7.Enabled = True     'FECHA DE NACIMIENTO
                Me.DateTimePicker1.Enabled = True  'DESDE
                Me.DateTimePicker2.Enabled = True  'HASTA
                Me.Button11.Enabled = True         'BOTON FECHAS
                '--------------- OTROS FILTROS
                'Me.ComboBox19.Enabled = True       'FUNCIONALIDAD DE EMPLEADO
                'Me.Button14.Enabled = True         'BOTON FUNCIONALIDAD DE EMPLEADO
                'Me.ComboBox20.Enabled = True       'CAUSA REAL DE EVENTUALIDAD
                'Me.Button15.Enabled = True         'BOTON CAUSA REAL DE EVENTUALIDAD
                'Me.ComboBox21.Enabled = True       'TIPO DE EVENTUALIDAD
                'Me.Button16.Enabled = True         'BOTON TIPO DE EVENTUALIDAD
                'Me.ComboBox22.Enabled = True       'ESTUDIOS
                'Me.Button17.Enabled = True         'BOTON ESTUDIOS
                'Me.ComboBox23.Enabled = True       'NIVEL DE COMPETENCIA
                'Me.Button18.Enabled = True         'BOTON NIVEL DE COMPETENCIA
                'Me.ComboBox26.Enabled = True       'CLASIFICACION DEL EMPLEADO
                'Me.Button21.Enabled = True         'BOTON CLASIFICACION
                Me.TextBox1.Enabled = True         'CODIGO DEL EMPLEADO
                Me.Button22.Enabled = True         'BOTON CODIGO EMPLEADO
                'Me.RadioButton12.Enabled = True    'EDAD
                'Me.RadioButton13.Enabled = True    'ANTIGUEDAD PAME
                'Me.RadioButton14.Enabled = True    'ANTIGUEDAD EN
                'Me.TextBox2.Enabled = True         'ANTIGUEDADES A
                'Me.TextBox3.Enabled = True         'ANTIGUEDADES B
                'Me.Button23.Enabled = True         'BOTON ANTIGUEDADES
                'Me.TextBox4.Enabled = True         'CONTRATOS
                'Me.Button24.Enabled = True         'BOTON CONTRATOS
                '--------------- CONDECORACIONES
                'Me.ComboBox27.Enabled = True      'CLASIFICACION
                'Me.Button25.Enabled = True        'BOTON CLASIFICACION
                'Me.ComboBox28.Enabled = True      'TIPO
                'Me.Button26.Enabled = True        'BOTON TIPO
                '--------------- SITUACIONES
                'Me.ComboBox29.Enabled = True      'TIPO SITUACION
                'Me.Button27.Enabled = True        'BOTON TIPO SITUACION
                'Me.ComboBox30.Enabled = True      'SITUACION
                'Me.Button28.Enabled = True        'BOTON SITUACION
                'Me.DateTimePicker3.Enabled = True 'DESDE
                'Me.DateTimePicker4.Enabled = True 'HASTA
                'Me.Button29.Enabled = True        'BOTON FECHAS

                ULTIMA_LISTA_QUE_HIZO_CLIC = Mid(CODIGO_REPORTE, 1, 1)
                GoTo SALIR
            End If
            If (CODIGO_REPORTE = "E0006") Then              'E0006 - Lista de Empleados por Funcionalidad
                '--------------- NIVELES ORGANIZACIONALES
                Me.ComboBox1.Enabled = True        'NIVEL 1
                Me.ComboBox2.Enabled = True        'NIVEL 2
                Me.CheckBox1.Enabled = True        'CHECK N2
                Me.ComboBox3.Enabled = True        'NIVEL 3
                Me.CheckBox2.Enabled = True        'CHECK N3
                Me.ComboBox4.Enabled = True        'NIVEL 4
                Me.CheckBox3.Enabled = True        'CHECK N4
                Me.ComboBox5.Enabled = True        'NIVEL 5
                Me.CheckBox4.Enabled = True        'CHECK N5
                Me.ComboBox6.Enabled = True        'NIVEL 6
                Me.CheckBox5.Enabled = True        'CHECK N6
                Me.ComboBox7.Enabled = True        'NIVEL 7
                Me.CheckBox6.Enabled = True        'CHECK N7
                Me.ComboBox8.Enabled = True        'NIVEL 8
                Me.CheckBox7.Enabled = True        'CHECK N8
                Me.Button1.Enabled = True          'BOTON NIVELES ORGANIZACIONALES
                '--------------- TIPO DE ESTRUCTURAS
                Me.RadioButton1.Enabled = True     'ESTRUCTURA MIXTA
                Me.RadioButton2.Enabled = True     'ESTRUCTURA MILITAR
                Me.RadioButton3.Enabled = True     'ESTRUCTURA PAME
                Me.RadioButton4.Enabled = True     'JEFE
                Me.Button2.Enabled = True          'BOTON TIPO DE ESTRUCTURAS
                '--------------- GRADOS, CARGOS, OTROS
                Me.ComboBox9.Enabled = True        'GRADO REAL
                Me.Button3.Enabled = True          'BOTON GR
                Me.ComboBox10.Enabled = True       'GRADO PLANTILLA
                Me.Button4.Enabled = True          'BOTON GP
                Me.ComboBox11.Enabled = True       'SITUACION DEL CARGO
                Me.Button5.Enabled = True          'BOTON SITUACION CARGO
                Me.ComboBox12.Enabled = True       'CARGO
                Me.Button6.Enabled = True          'BOTON CARGO
                Me.ComboBox13.Enabled = True       'ESTABLECIMIENTO
                Me.Button7.Enabled = True          'BOTON ESTABLECIMIENTO
                '--------------- FECHAS
                Me.RadioButton5.Enabled = True     'FECHA DE INGRESO PAME
                Me.RadioButton5.Checked = True
                Me.RadioButton6.Enabled = True     'FECHA DE INGRESO EN
                Me.RadioButton7.Enabled = True     'FECHA DE NACIMIENTO
                Me.DateTimePicker1.Enabled = True  'DESDE
                Me.DateTimePicker2.Enabled = True  'HASTA
                Me.Button11.Enabled = True         'BOTON FECHAS
                '--------------- OTROS FILTROS
                Me.ComboBox19.Enabled = True       'FUNCIONALIDAD DE EMPLEADO
                Me.Button14.Enabled = True         'BOTON FUNCIONALIDAD DE EMPLEADO
                'Me.ComboBox20.Enabled = True       'CAUSA REAL DE EVENTUALIDAD
                'Me.Button15.Enabled = True         'BOTON CAUSA REAL DE EVENTUALIDAD
                'Me.ComboBox21.Enabled = True       'TIPO DE EVENTUALIDAD
                'Me.Button16.Enabled = True         'BOTON TIPO DE EVENTUALIDAD
                'Me.ComboBox22.Enabled = True       'ESTUDIOS
                'Me.Button17.Enabled = True         'BOTON ESTUDIOS
                Me.ComboBox23.Enabled = True       'NIVEL DE COMPETENCIA
                Me.Button18.Enabled = True         'BOTON NIVEL DE COMPETENCIA
                'Me.ComboBox26.Enabled = True       'CLASIFICACION DEL EMPLEADO
                'Me.Button21.Enabled = True         'BOTON CLASIFICACION
                Me.TextBox1.Enabled = True         'CODIGO DEL EMPLEADO
                Me.Button22.Enabled = True         'BOTON CODIGO EMPLEADO
                'Me.RadioButton12.Enabled = True    'EDAD
                'Me.RadioButton13.Enabled = True    'ANTIGUEDAD PAME
                'Me.RadioButton14.Enabled = True    'ANTIGUEDAD EN
                'Me.TextBox2.Enabled = True         'ANTIGUEDADES A
                'Me.TextBox3.Enabled = True         'ANTIGUEDADES B
                'Me.Button23.Enabled = True         'BOTON ANTIGUEDADES
                'Me.TextBox4.Enabled = True         'CONTRATOS
                'Me.Button24.Enabled = True         'BOTON CONTRATOS
                '--------------- CONDECORACIONES
                'Me.ComboBox27.Enabled = True      'CLASIFICACION
                'Me.Button25.Enabled = True        'BOTON CLASIFICACION
                'Me.ComboBox28.Enabled = True      'TIPO
                'Me.Button26.Enabled = True        'BOTON TIPO
                '--------------- SITUACIONES
                'Me.ComboBox29.Enabled = True      'TIPO SITUACION
                'Me.Button27.Enabled = True        'BOTON TIPO SITUACION
                'Me.ComboBox30.Enabled = True      'SITUACION
                'Me.Button28.Enabled = True        'BOTON SITUACION
                'Me.DateTimePicker3.Enabled = True 'DESDE
                'Me.DateTimePicker4.Enabled = True 'HASTA
                'Me.Button29.Enabled = True        'BOTON FECHAS

                ULTIMA_LISTA_QUE_HIZO_CLIC = Mid(CODIGO_REPORTE, 1, 1)
                GoTo SALIR
            End If
            If (CODIGO_REPORTE = "E0007") Then                      'E0007 - Contratos y Evaluaciones por Empleado
                '--------------- NIVELES ORGANIZACIONALES
                Me.ComboBox1.Enabled = True        'NIVEL 1
                Me.ComboBox2.Enabled = True        'NIVEL 2
                Me.CheckBox1.Enabled = True        'CHECK N2
                Me.ComboBox3.Enabled = True        'NIVEL 3
                Me.CheckBox2.Enabled = True        'CHECK N3
                Me.ComboBox4.Enabled = True        'NIVEL 4
                Me.CheckBox3.Enabled = True        'CHECK N4
                Me.ComboBox5.Enabled = True        'NIVEL 5
                Me.CheckBox4.Enabled = True        'CHECK N5
                Me.ComboBox6.Enabled = True        'NIVEL 6
                Me.CheckBox5.Enabled = True        'CHECK N6
                Me.ComboBox7.Enabled = True        'NIVEL 7
                Me.CheckBox6.Enabled = True        'CHECK N7
                Me.ComboBox8.Enabled = True        'NIVEL 8
                Me.CheckBox7.Enabled = True        'CHECK N8
                Me.Button1.Enabled = True          'BOTON NIVELES ORGANIZACIONALES
                '--------------- TIPO DE ESTRUCTURAS
                Me.RadioButton1.Enabled = True     'ESTRUCTURA MIXTA
                Me.RadioButton2.Enabled = True     'ESTRUCTURA MILITAR
                Me.RadioButton3.Enabled = True     'ESTRUCTURA PAME
                Me.RadioButton4.Enabled = True     'JEFE
                Me.Button2.Enabled = True          'BOTON TIPO DE ESTRUCTURAS
                '--------------- GRADOS, CARGOS, OTROS
                Me.ComboBox9.Enabled = True        'GRADO REAL
                Me.Button3.Enabled = True          'BOTON GR
                Me.ComboBox10.Enabled = True       'GRADO PLANTILLA
                Me.Button4.Enabled = True          'BOTON GP
                Me.ComboBox11.Enabled = True       'SITUACION DEL CARGO
                Me.Button5.Enabled = True          'BOTON SITUACION CARGO
                Me.ComboBox12.Enabled = True       'CARGO
                Me.Button6.Enabled = True          'BOTON CARGO
                Me.ComboBox13.Enabled = True       'ESTABLECIMIENTO
                Me.Button7.Enabled = True          'BOTON ESTABLECIMIENTO
                '--------------- FECHAS
                Me.RadioButton5.Enabled = True     'FECHA DE INGRESO PAME
                Me.RadioButton5.Checked = True
                Me.RadioButton6.Enabled = True     'FECHA DE INGRESO EN
                Me.RadioButton7.Enabled = True     'FECHA DE NACIMIENTO
                Me.DateTimePicker1.Enabled = True  'DESDE
                Me.DateTimePicker2.Enabled = True  'HASTA
                Me.Button11.Enabled = True         'BOTON FECHAS
                '--------------- OTROS FILTROS
                'Me.ComboBox19.Enabled = True       'FUNCIONALIDAD DE EMPLEADO
                'Me.Button14.Enabled = True         'BOTON FUNCIONALIDAD DE EMPLEADO
                'Me.ComboBox20.Enabled = True       'CAUSA REAL DE EVENTUALIDAD
                'Me.Button15.Enabled = True         'BOTON CAUSA REAL DE EVENTUALIDAD
                'Me.ComboBox21.Enabled = True       'TIPO DE EVENTUALIDAD
                'Me.Button16.Enabled = True         'BOTON TIPO DE EVENTUALIDAD
                'Me.ComboBox22.Enabled = True       'ESTUDIOS
                'Me.Button17.Enabled = True         'BOTON ESTUDIOS
                'Me.ComboBox23.Enabled = True       'NIVEL DE COMPETENCIA
                'Me.Button18.Enabled = True         'BOTON NIVEL DE COMPETENCIA
                'Me.ComboBox26.Enabled = True       'CLASIFICACION DEL EMPLEADO
                'Me.Button21.Enabled = True         'BOTON CLASIFICACION
                Me.TextBox1.Enabled = True         'CODIGO DEL EMPLEADO
                Me.Button22.Enabled = True         'BOTON CODIGO EMPLEADO
                'Me.RadioButton12.Enabled = True    'EDAD
                'Me.RadioButton13.Enabled = True    'ANTIGUEDAD PAME
                'Me.RadioButton14.Enabled = True    'ANTIGUEDAD EN
                'Me.TextBox2.Enabled = True         'ANTIGUEDADES A
                'Me.TextBox3.Enabled = True         'ANTIGUEDADES B
                'Me.Button23.Enabled = True         'BOTON ANTIGUEDADES
                Me.TextBox4.Enabled = True         'CONTRATOS
                Me.Button24.Enabled = True         'BOTON CONTRATOS
                '--------------- CONDECORACIONES
                'Me.ComboBox27.Enabled = True      'CLASIFICACION
                'Me.Button25.Enabled = True        'BOTON CLASIFICACION
                'Me.ComboBox28.Enabled = True      'TIPO
                'Me.Button26.Enabled = True        'BOTON TIPO
                '--------------- SITUACIONES
                'Me.ComboBox29.Enabled = True      'TIPO SITUACION
                'Me.Button27.Enabled = True        'BOTON TIPO SITUACION
                'Me.ComboBox30.Enabled = True      'SITUACION
                'Me.Button28.Enabled = True        'BOTON SITUACION
                'Me.DateTimePicker3.Enabled = True 'DESDE
                'Me.DateTimePicker4.Enabled = True 'HASTA
                'Me.Button29.Enabled = True        'BOTON FECHAS

                ULTIMA_LISTA_QUE_HIZO_CLIC = Mid(CODIGO_REPORTE, 1, 1)
                GoTo SALIR
            End If
            If (CODIGO_REPORTE = "E0008" Or CODIGO_REPORTE = "E0013") Then                      'E0008 - Eventualidades, E0013 - Aplicacion del reglamento disciplinario
                '--------------- NIVELES ORGANIZACIONALES
                Me.ComboBox1.Enabled = True        'NIVEL 1
                Me.ComboBox2.Enabled = True        'NIVEL 2
                Me.CheckBox1.Enabled = True        'CHECK N2
                Me.ComboBox3.Enabled = True        'NIVEL 3
                Me.CheckBox2.Enabled = True        'CHECK N3
                Me.ComboBox4.Enabled = True        'NIVEL 4
                Me.CheckBox3.Enabled = True        'CHECK N4
                Me.ComboBox5.Enabled = True        'NIVEL 5
                Me.CheckBox4.Enabled = True        'CHECK N5
                Me.ComboBox6.Enabled = True        'NIVEL 6
                Me.CheckBox5.Enabled = True        'CHECK N6
                Me.ComboBox7.Enabled = True        'NIVEL 7
                Me.CheckBox6.Enabled = True        'CHECK N7
                Me.ComboBox8.Enabled = True        'NIVEL 8
                Me.CheckBox7.Enabled = True        'CHECK N8
                Me.Button1.Enabled = True          'BOTON NIVELES ORGANIZACIONALES
                '--------------- TIPO DE ESTRUCTURAS
                Me.RadioButton1.Enabled = True     'ESTRUCTURA MIXTA
                Me.RadioButton2.Enabled = True     'ESTRUCTURA MILITAR
                Me.RadioButton3.Enabled = True     'ESTRUCTURA PAME
                Me.RadioButton4.Enabled = True     'JEFE
                Me.Button2.Enabled = True          'BOTON TIPO DE ESTRUCTURAS
                '--------------- GRADOS, CARGOS, OTROS
                Me.ComboBox9.Enabled = True        'GRADO REAL
                Me.Button3.Enabled = True          'BOTON GR
                Me.ComboBox10.Enabled = True       'GRADO PLANTILLA
                Me.Button4.Enabled = True          'BOTON GP
                Me.ComboBox11.Enabled = True       'SITUACION DEL CARGO
                Me.Button5.Enabled = True          'BOTON SITUACION CARGO
                Me.ComboBox12.Enabled = True       'CARGO
                Me.Button6.Enabled = True          'BOTON CARGO
                Me.ComboBox13.Enabled = True       'ESTABLECIMIENTO
                Me.Button7.Enabled = True          'BOTON ESTABLECIMIENTO
                '--------------- FECHAS
                Me.RadioButton5.Enabled = True     'FECHA DE INGRESO PAME
                Me.RadioButton5.Checked = True
                Me.RadioButton6.Enabled = True     'FECHA DE INGRESO EN
                Me.RadioButton7.Enabled = True     'FECHA DE NACIMIENTO
                Me.RadioButton15.Enabled = True    'FECHA DE EVENTUALIDADES
                'Me.RadioButton16.Enabled = TRUE   'FECHA DE SUBSIDIOS
                Me.DateTimePicker1.Enabled = True  'DESDE
                Me.DateTimePicker2.Enabled = True  'HASTA
                Me.Button11.Enabled = True         'BOTON FECHAS
                '--------------- OTROS FILTROS
                Me.ComboBox19.Enabled = True       'FUNCIONALIDAD DE EMPLEADO
                Me.Button14.Enabled = True         'BOTON FUNCIONALIDAD DE EMPLEADO
                Me.ComboBox20.Enabled = True       'CAUSA REAL DE EVENTUALIDAD
                Me.Button15.Enabled = True         'BOTON CAUSA REAL DE EVENTUALIDAD
                Me.ComboBox21.Enabled = True       'TIPO DE EVENTUALIDAD
                Me.Button16.Enabled = True         'BOTON TIPO DE EVENTUALIDAD
                'Me.ComboBox22.Enabled = True       'ESTUDIOS
                'Me.Button17.Enabled = True         'BOTON ESTUDIOS
                Me.ComboBox23.Enabled = True       'NIVEL DE COMPETENCIA
                Me.Button18.Enabled = True         'BOTON NIVEL DE COMPETENCIA
                'Me.ComboBox26.Enabled = True       'CLASIFICACION DEL EMPLEADO
                'Me.Button21.Enabled = True         'BOTON CLASIFICACION
                Me.TextBox1.Enabled = True         'CODIGO DEL EMPLEADO
                Me.Button22.Enabled = True         'BOTON CODIGO EMPLEADO
                'Me.RadioButton12.Enabled = True    'EDAD
                'Me.RadioButton13.Enabled = True    'ANTIGUEDAD PAME
                'Me.RadioButton14.Enabled = True    'ANTIGUEDAD EN
                'Me.TextBox2.Enabled = True         'ANTIGUEDADES A
                'Me.TextBox3.Enabled = True         'ANTIGUEDADES B
                'Me.Button23.Enabled = True         'BOTON ANTIGUEDADES
                'Me.TextBox4.Enabled = True         'CONTRATOS
                'Me.Button24.Enabled = True         'BOTON CONTRATOS
                '--------------- CONDECORACIONES
                'Me.ComboBox27.Enabled = True      'CLASIFICACION
                'Me.Button25.Enabled = True        'BOTON CLASIFICACION
                'Me.ComboBox28.Enabled = True      'TIPO
                'Me.Button26.Enabled = True        'BOTON TIPO
                '--------------- SITUACIONES
                'Me.ComboBox29.Enabled = True      'TIPO SITUACION
                'Me.Button27.Enabled = True        'BOTON TIPO SITUACION
                'Me.ComboBox30.Enabled = True      'SITUACION
                'Me.Button28.Enabled = True        'BOTON SITUACION
                'Me.DateTimePicker3.Enabled = True 'DESDE
                'Me.DateTimePicker4.Enabled = True 'HASTA
                'Me.Button29.Enabled = True        'BOTON FECHAS

                ULTIMA_LISTA_QUE_HIZO_CLIC = Mid(CODIGO_REPORTE, 1, 1)
                GoTo SALIR
            End If

            If (CODIGO_REPORTE = "E0009" Or CODIGO_REPORTE = "E0010" Or CODIGO_REPORTE = "E0011" Or CODIGO_REPORTE = "E0012") Then              'CONSTANCIAS
                Frm_Reportes_Constancias.ShowDialog()
                ULTIMA_LISTA_QUE_HIZO_CLIC = Mid(CODIGO_REPORTE, 1, 1)
                GoTo SALIR
            End If

SALIR:
        End If
    End Sub
    Private Sub ListBox5_Click(sender As Object, e As EventArgs) Handles ListBox5.Click
        If Me.ListBox5.Text <> "" Then
            Call DESACTIVAR_TODO()
            CODIGO_REPORTE = Mid(Me.ListBox5.Text, 1, 5)
            NOMBRE_REPORTE = Me.ListBox5.Text
            If Mid(CODIGO_REPORTE, 1, 1) <> ULTIMA_LISTA_QUE_HIZO_CLIC Then
                Call LIMPIAR_LISTAS_DE_SELECCION()
                Me.Lista_Principal.Items.Clear()
            End If
            Label2.Text = NOMBRE_REPORTE
            Label2.Text = "Medicos Acreditados\" & NOMBRE_REPORTE
        End If
    End Sub
    Private Sub ListBox6_Click(sender As Object, e As EventArgs) Handles ListBox6.Click
        If Me.ListBox6.Text <> "" Then
            Call DESACTIVAR_TODO()
            CODIGO_REPORTE = Mid(Me.ListBox6.Text, 1, 5)
            If Mid(CODIGO_REPORTE, 1, 1) <> ULTIMA_LISTA_QUE_HIZO_CLIC Then
                Call LIMPIAR_LISTAS_DE_SELECCION()
                Me.Lista_Principal.Items.Clear()
            End If
            NOMBRE_REPORTE = Me.ListBox6.Text
            Label2.Text = NOMBRE_REPORTE
            Label2.Text = "Becados\" & NOMBRE_REPORTE
            If (CODIGO_REPORTE = "B0002") Then  'CONSTANCIAS BECADOS

            End If

            ULTIMA_LISTA_QUE_HIZO_CLIC = Mid(CODIGO_REPORTE, 1, 1)
SALIR:
        End If
    End Sub
    Private Sub ListBox7_Click(sender As Object, e As EventArgs) Handles ListBox7.Click
        If Me.ListBox7.Text <> "" Then
            Call DESACTIVAR_TODO()
            CODIGO_REPORTE = Mid(Me.ListBox7.Text, 1, 5)
            If Mid(CODIGO_REPORTE, 1, 1) <> ULTIMA_LISTA_QUE_HIZO_CLIC Then
                Call LIMPIAR_LISTAS_DE_SELECCION()
                Me.Lista_Principal.Items.Clear()
            End If
            NOMBRE_REPORTE = Me.ListBox7.Text
            Label2.Text = NOMBRE_REPORTE
            Label2.Text = "Asistencia y Situaciones del Personal\" & NOMBRE_REPORTE
            If (CODIGO_REPORTE = "S0001") Then                      'S0001 - Marcas por Empleado (ACTIVOS - INACTIVOS)
                '--------------- NIVELES ORGANIZACIONALES
                'Me.ComboBox1.Enabled = True        'NIVEL 1
                'Me.ComboBox2.Enabled = True        'NIVEL 2
                'Me.CheckBox1.Enabled = True        'CHECK N2
                'Me.ComboBox3.Enabled = True        'NIVEL 3
                'Me.CheckBox2.Enabled = True        'CHECK N3
                'Me.ComboBox4.Enabled = True        'NIVEL 4
                'Me.CheckBox3.Enabled = True        'CHECK N4
                'Me.ComboBox5.Enabled = True        'NIVEL 5
                'Me.CheckBox4.Enabled = True        'CHECK N5
                'Me.ComboBox6.Enabled = True        'NIVEL 6
                'Me.CheckBox5.Enabled = True        'CHECK N6
                'Me.ComboBox7.Enabled = True        'NIVEL 7
                'Me.CheckBox6.Enabled = True        'CHECK N7
                'Me.ComboBox8.Enabled = True        'NIVEL 8
                'Me.CheckBox7.Enabled = True        'CHECK N8
                'Me.Button1.Enabled = True          'BOTON NIVELES ORGANIZACIONALES
                '--------------- TIPO DE ESTRUCTURAS
                'Me.RadioButton1.Enabled = True     'ESTRUCTURA MIXTA
                'Me.RadioButton2.Enabled = True     'ESTRUCTURA MILITAR
                'Me.RadioButton3.Enabled = True     'ESTRUCTURA PAME
                'Me.RadioButton4.Enabled = True     'JEFE
                'Me.Button2.Enabled = True          'BOTON TIPO DE ESTRUCTURAS
                '--------------- GRADOS, CARGOS, OTROS
                'Me.ComboBox9.Enabled = True        'GRADO REAL
                'Me.Button3.Enabled = True          'BOTON GR
                'Me.ComboBox10.Enabled = True       'GRADO PLANTILLA
                'Me.Button4.Enabled = True          'BOTON GP
                'Me.ComboBox11.Enabled = True       'SITUACION DEL CARGO
                'Me.Button5.Enabled = True          'BOTON SITUACION CARGO
                'Me.ComboBox12.Enabled = True       'CARGO
                'Me.Button6.Enabled = True          'BOTON CARGO
                'Me.ComboBox13.Enabled = True       'ESTABLECIMIENTO
                'Me.Button7.Enabled = True          'BOTON ESTABLECIMIENTO
                '--------------- FECHAS
                'Me.RadioButton5.Enabled = True     'FECHA DE INGRESO PAME
                'Me.RadioButton5.Checked = True
                'Me.RadioButton6.Enabled = True     'FECHA DE INGRESO EN
                'Me.RadioButton7.Enabled = True     'FECHA DE NACIMIENTO
                'Me.RadioButton8.Enabled = True     'FECHA DE SITUACION
                Me.RadioButton9.Enabled = True      'FECHA DE MARCA
                Me.RadioButton9.Checked = True
                'Me.RadioButton10.Enabled = True    'FECHA DE PROCESO DEL ASPIRANTE
                'Me.RadioButton11.Enabled = True    'FECHA DE APLICACION DE MOVIMIENTOS
                'Me.RadioButton15.Enabled = True    'FECHA DE EVENTUALIDADES
                'Me.RadioButton16.Enabled = TRUE    'FECHA DE SUBSIDIOS
                Me.DateTimePicker1.Enabled = True   'DESDE
                Me.DateTimePicker2.Enabled = True   'HASTA
                Me.Button11.Enabled = True          'BOTON FECHAS
                '--------------- OTROS FILTROS
                'Me.ComboBox19.Enabled = True       'FUNCIONALIDAD DE EMPLEADO
                'Me.Button14.Enabled = True         'BOTON FUNCIONALIDAD DE EMPLEADO
                'Me.ComboBox20.Enabled = True       'CAUSA REAL DE EVENTUALIDAD
                'Me.Button15.Enabled = True         'BOTON CAUSA REAL DE EVENTUALIDAD
                'Me.ComboBox21.Enabled = True       'TIPO DE EVENTUALIDAD
                'Me.Button16.Enabled = True         'BOTON TIPO DE EVENTUALIDAD
                'Me.ComboBox22.Enabled = True       'ESTUDIOS
                'Me.Button17.Enabled = True         'BOTON ESTUDIOS
                'Me.ComboBox23.Enabled = True       'NIVEL DE COMPETENCIA
                'Me.Button18.Enabled = True         'BOTON NIVEL DE COMPETENCIA
                Me.ComboBox24.Enabled = True        'RELOJ DE MARCACION
                Me.Button19.Enabled = True          'BOTON RELOJ DE MARCACION
                Me.ComboBox25.Enabled = True        'TIPO DE MARCA
                Me.Button20.Enabled = True          'BOTON TIPO DE MARCA
                'Me.ComboBox26.Enabled = True       'CLASIFICACION DEL EMPLEADO
                'Me.Button21.Enabled = True         'BOTON CLASIFICACION
                Me.TextBox1.Enabled = True          'CODIGO DEL EMPLEADO
                Me.Button22.Enabled = True          'BOTON CODIGO EMPLEADO
                'Me.RadioButton12.Enabled = True    'EDAD
                'Me.RadioButton13.Enabled = True    'ANTIGUEDAD PAME
                'Me.RadioButton14.Enabled = True    'ANTIGUEDAD EN
                'Me.TextBox2.Enabled = True         'ANTIGUEDADES A
                'Me.TextBox3.Enabled = True         'ANTIGUEDADES B
                'Me.Button23.Enabled = True         'BOTON ANTIGUEDADES
                'Me.TextBox4.Enabled = True         'CONTRATOS
                'Me.Button24.Enabled = True         'BOTON CONTRATOS
                '--------------- CONDECORACIONES
                'Me.ComboBox27.Enabled = True      'CLASIFICACION
                'Me.Button25.Enabled = True        'BOTON CLASIFICACION
                'Me.ComboBox28.Enabled = True      'TIPO
                'Me.Button26.Enabled = True        'BOTON TIPO
                '--------------- SITUACIONES
                'Me.ComboBox29.Enabled = True      'TIPO SITUACION
                'Me.Button27.Enabled = True        'BOTON TIPO SITUACION
                'Me.ComboBox30.Enabled = True      'SITUACION
                'Me.Button28.Enabled = True        'BOTON SITUACION
                'Me.DateTimePicker3.Enabled = True 'DESDE
                'Me.DateTimePicker4.Enabled = True 'HASTA
                'Me.Button29.Enabled = True        'BOTON FECHAS

                ULTIMA_LISTA_QUE_HIZO_CLIC = Mid(CODIGO_REPORTE, 1, 1)
                GoTo SALIR
            End If
            If (CODIGO_REPORTE = "S0002") Then                      'S0002 - Marcas por Niveles Organizacionales
                '--------------- NIVELES ORGANIZACIONALES
                Me.ComboBox1.Enabled = True        'NIVEL 1
                Me.ComboBox2.Enabled = True        'NIVEL 2
                Me.CheckBox1.Enabled = True        'CHECK N2
                Me.ComboBox3.Enabled = True        'NIVEL 3
                Me.CheckBox2.Enabled = True        'CHECK N3
                Me.ComboBox4.Enabled = True        'NIVEL 4
                Me.CheckBox3.Enabled = True        'CHECK N4
                Me.ComboBox5.Enabled = True        'NIVEL 5
                Me.CheckBox4.Enabled = True        'CHECK N5
                Me.ComboBox6.Enabled = True        'NIVEL 6
                Me.CheckBox5.Enabled = True        'CHECK N6
                Me.ComboBox7.Enabled = True        'NIVEL 7
                Me.CheckBox6.Enabled = True        'CHECK N7
                Me.ComboBox8.Enabled = True        'NIVEL 8
                Me.CheckBox7.Enabled = True        'CHECK N8
                Me.Button1.Enabled = True          'BOTON NIVELES ORGANIZACIONALES
                '--------------- TIPO DE ESTRUCTURAS
                'Me.RadioButton1.Enabled = True     'ESTRUCTURA MIXTA
                'Me.RadioButton2.Enabled = True     'ESTRUCTURA MILITAR
                'Me.RadioButton3.Enabled = True     'ESTRUCTURA PAME
                'Me.RadioButton4.Enabled = True     'JEFE
                'Me.Button2.Enabled = True          'BOTON TIPO DE ESTRUCTURAS
                '--------------- GRADOS, CARGOS, OTROS
                'Me.ComboBox9.Enabled = True        'GRADO REAL
                'Me.Button3.Enabled = True          'BOTON GR
                'Me.ComboBox10.Enabled = True       'GRADO PLANTILLA
                'Me.Button4.Enabled = True          'BOTON GP
                'Me.ComboBox11.Enabled = True       'SITUACION DEL CARGO
                'Me.Button5.Enabled = True          'BOTON SITUACION CARGO
                'Me.ComboBox12.Enabled = True       'CARGO
                'Me.Button6.Enabled = True          'BOTON CARGO
                'Me.ComboBox13.Enabled = True       'ESTABLECIMIENTO
                'Me.Button7.Enabled = True          'BOTON ESTABLECIMIENTO
                '--------------- FECHAS
                'Me.RadioButton5.Enabled = True     'FECHA DE INGRESO PAME
                'Me.RadioButton5.Checked = True
                'Me.RadioButton6.Enabled = True     'FECHA DE INGRESO EN
                'Me.RadioButton7.Enabled = True     'FECHA DE NACIMIENTO
                'Me.RadioButton8.Enabled = True     'FECHA DE SITUACION
                Me.RadioButton9.Enabled = True      'FECHA DE MARCA
                Me.RadioButton9.Checked = True
                'Me.RadioButton10.Enabled = True    'FECHA DE PROCESO DEL ASPIRANTE
                'Me.RadioButton11.Enabled = True    'FECHA DE APLICACION DE MOVIMIENTOS
                'Me.RadioButton15.Enabled = True    'FECHA DE EVENTUALIDADES
                'Me.RadioButton16.Enabled = TRUE    'FECHA DE SUBSIDIOS
                Me.DateTimePicker1.Enabled = True   'DESDE
                Me.DateTimePicker2.Enabled = True   'HASTA
                Me.Button11.Enabled = True          'BOTON FECHAS
                '--------------- OTROS FILTROS
                'Me.ComboBox19.Enabled = True       'FUNCIONALIDAD DE EMPLEADO
                'Me.Button14.Enabled = True         'BOTON FUNCIONALIDAD DE EMPLEADO
                'Me.ComboBox20.Enabled = True       'CAUSA REAL DE EVENTUALIDAD
                'Me.Button15.Enabled = True         'BOTON CAUSA REAL DE EVENTUALIDAD
                'Me.ComboBox21.Enabled = True       'TIPO DE EVENTUALIDAD
                'Me.Button16.Enabled = True         'BOTON TIPO DE EVENTUALIDAD
                'Me.ComboBox22.Enabled = True       'ESTUDIOS
                'Me.Button17.Enabled = True         'BOTON ESTUDIOS
                'Me.ComboBox23.Enabled = True       'NIVEL DE COMPETENCIA
                'Me.Button18.Enabled = True         'BOTON NIVEL DE COMPETENCIA
                Me.ComboBox24.Enabled = True        'RELOJ DE MARCACION
                Me.Button19.Enabled = True          'BOTON RELOJ DE MARCACION
                Me.ComboBox25.Enabled = True        'TIPO DE MARCA
                Me.Button20.Enabled = True          'BOTON TIPO DE MARCA
                'Me.ComboBox26.Enabled = True       'CLASIFICACION DEL EMPLEADO
                'Me.Button21.Enabled = True         'BOTON CLASIFICACION
                Me.TextBox1.Enabled = True          'CODIGO DEL EMPLEADO
                Me.Button22.Enabled = True          'BOTON CODIGO EMPLEADO
                'Me.RadioButton12.Enabled = True    'EDAD
                'Me.RadioButton13.Enabled = True    'ANTIGUEDAD PAME
                'Me.RadioButton14.Enabled = True    'ANTIGUEDAD EN
                'Me.TextBox2.Enabled = True         'ANTIGUEDADES A
                'Me.TextBox3.Enabled = True         'ANTIGUEDADES B
                'Me.Button23.Enabled = True         'BOTON ANTIGUEDADES
                'Me.TextBox4.Enabled = True         'CONTRATOS
                'Me.Button24.Enabled = True         'BOTON CONTRATOS
                '--------------- CONDECORACIONES
                'Me.ComboBox27.Enabled = True      'CLASIFICACION
                'Me.Button25.Enabled = True        'BOTON CLASIFICACION
                'Me.ComboBox28.Enabled = True      'TIPO
                'Me.Button26.Enabled = True        'BOTON TIPO
                '--------------- SITUACIONES
                'Me.ComboBox29.Enabled = True      'TIPO SITUACION
                'Me.Button27.Enabled = True        'BOTON TIPO SITUACION
                'Me.ComboBox30.Enabled = True      'SITUACION
                'Me.Button28.Enabled = True        'BOTON SITUACION
                'Me.DateTimePicker3.Enabled = True 'DESDE
                'Me.DateTimePicker4.Enabled = True 'HASTA
                'Me.Button29.Enabled = True        'BOTON FECHAS

                ULTIMA_LISTA_QUE_HIZO_CLIC = Mid(CODIGO_REPORTE, 1, 1)
                GoTo SALIR
            End If
            If (CODIGO_REPORTE = "S0003") Then                      'S0003 - Formato de Solicitud de Vacaciones
                Frm_Formato_Vacaciones.ShowDialog()
                ULTIMA_LISTA_QUE_HIZO_CLIC = Mid(CODIGO_REPORTE, 1, 1)
                GoTo SALIR
            End If
            If (CODIGO_REPORTE = "S0004") Then                      'S0004 - Cumpleañeros
                '--------------- NIVELES ORGANIZACIONALES
                Me.ComboBox1.Enabled = True        'NIVEL 1
                Me.ComboBox2.Enabled = True        'NIVEL 2
                Me.CheckBox1.Enabled = True        'CHECK N2
                Me.ComboBox3.Enabled = True        'NIVEL 3
                Me.CheckBox2.Enabled = True        'CHECK N3
                Me.ComboBox4.Enabled = True        'NIVEL 4
                Me.CheckBox3.Enabled = True        'CHECK N4
                Me.ComboBox5.Enabled = True        'NIVEL 5
                Me.CheckBox4.Enabled = True        'CHECK N5
                Me.ComboBox6.Enabled = True        'NIVEL 6
                Me.CheckBox5.Enabled = True        'CHECK N6
                Me.ComboBox7.Enabled = True        'NIVEL 7
                Me.CheckBox6.Enabled = True        'CHECK N7
                Me.ComboBox8.Enabled = True        'NIVEL 8
                Me.CheckBox7.Enabled = True        'CHECK N8
                Me.Button1.Enabled = True          'BOTON NIVELES ORGANIZACIONALES
                '--------------- TIPO DE ESTRUCTURAS
                Me.RadioButton1.Enabled = True     'ESTRUCTURA MIXTA
                Me.RadioButton2.Enabled = True     'ESTRUCTURA MILITAR
                Me.RadioButton3.Enabled = True     'ESTRUCTURA PAME
                Me.RadioButton4.Enabled = True     'JEFE
                Me.Button2.Enabled = True          'BOTON TIPO DE ESTRUCTURAS
                '--------------- GRADOS, CARGOS, OTROS
                Me.ComboBox9.Enabled = True        'GRADO REAL
                Me.Button3.Enabled = True          'BOTON GR
                Me.ComboBox10.Enabled = True       'GRADO PLANTILLA
                Me.Button4.Enabled = True          'BOTON GP
                Me.ComboBox11.Enabled = True       'SITUACION DEL CARGO
                Me.Button5.Enabled = True          'BOTON SITUACION CARGO
                Me.ComboBox12.Enabled = True       'CARGO
                Me.Button6.Enabled = True          'BOTON CARGO
                Me.ComboBox13.Enabled = True       'ESTABLECIMIENTO
                Me.Button7.Enabled = True          'BOTON ESTABLECIMIENTO
                '--------------- FECHAS
                'Me.RadioButton5.Enabled = True    'FECHA DE INGRESO PAME
                'Me.RadioButton5.Checked = True
                'Me.RadioButton6.Enabled = True    'FECHA DE INGRESO EN
                Me.RadioButton7.Enabled = True     'FECHA DE NACIMIENTO
                Me.RadioButton7.Checked = True
                'Me.RadioButton8.Enabled = True    'FECHA DE SITUACION
                'Me.RadioButton9.Enabled = True    'FECHA DE MARCA
                'Me.RadioButton9.Checked = True
                'Me.RadioButton10.Enabled = True   'FECHA DE PROCESO DEL ASPIRANTE
                'Me.RadioButton11.Enabled = True   'FECHA DE APLICACION DE MOVIMIENTOS
                'Me.RadioButton15.Enabled = True   'FECHA DE EVENTUALIDADES
                'Me.RadioButton16.Enabled = TRUE    'FECHA DE SUBSIDIOS
                Me.DateTimePicker1.Enabled = True  'DESDE
                Me.DateTimePicker2.Enabled = True  'HASTA
                Me.Button11.Enabled = True         'BOTON FECHAS
                '--------------- OTROS FILTROS
                'Me.ComboBox19.Enabled = True      'FUNCIONALIDAD DE EMPLEADO
                'Me.Button14.Enabled = True        'BOTON FUNCIONALIDAD DE EMPLEADO
                'Me.ComboBox20.Enabled = True      'CAUSA REAL DE EVENTUALIDAD
                'Me.Button15.Enabled = True        'BOTON CAUSA REAL DE EVENTUALIDAD
                'Me.ComboBox21.Enabled = True      'TIPO DE EVENTUALIDAD
                'Me.Button16.Enabled = True        'BOTON TIPO DE EVENTUALIDAD
                'Me.ComboBox22.Enabled = True      'ESTUDIOS
                'Me.Button17.Enabled = True        'BOTON ESTUDIOS
                'Me.ComboBox23.Enabled = True      'NIVEL DE COMPETENCIA
                'Me.Button18.Enabled = True        'BOTON NIVEL DE COMPETENCIA
                'Me.ComboBox24.Enabled = True      'RELOJ DE MARCACION
                'Me.Button19.Enabled = True        'BOTON RELOJ DE MARCACION
                'Me.ComboBox25.Enabled = True      'TIPO DE MARCA
                'Me.Button20.Enabled = True        'BOTON TIPO DE MARCA
                'Me.ComboBox26.Enabled = True      'CLASIFICACION DEL EMPLEADO
                'Me.Button21.Enabled = True        'BOTON CLASIFICACION
                'Me.TextBox1.Enabled = True        'CODIGO DEL EMPLEADO
                'Me.Button22.Enabled = True        'BOTON CODIGO EMPLEADO
                Me.RadioButton12.Enabled = True    'EDAD
                'Me.RadioButton13.Enabled = True   'ANTIGUEDAD PAME
                'Me.RadioButton14.Enabled = True   'ANTIGUEDAD EN
                Me.TextBox2.Enabled = True         'ANTIGUEDADES A
                Me.TextBox3.Enabled = True         'ANTIGUEDADES B
                Me.Button23.Enabled = True         'BOTON ANTIGUEDADES
                'Me.TextBox4.Enabled = True        'CONTRATOS
                'Me.Button24.Enabled = True        'BOTON CONTRATOS
                '--------------- CONDECORACIONES
                'Me.ComboBox27.Enabled = True      'CLASIFICACION
                'Me.Button25.Enabled = True        'BOTON CLASIFICACION
                'Me.ComboBox28.Enabled = True      'TIPO
                'Me.Button26.Enabled = True        'BOTON TIPO
                '--------------- SITUACIONES
                'Me.ComboBox29.Enabled = True      'TIPO SITUACION
                'Me.Button27.Enabled = True        'BOTON TIPO SITUACION
                'Me.ComboBox30.Enabled = True      'SITUACION
                'Me.Button28.Enabled = True        'BOTON SITUACION
                'Me.DateTimePicker3.Enabled = True 'DESDE
                'Me.DateTimePicker4.Enabled = True 'HASTA
                'Me.Button29.Enabled = True        'BOTON FECHAS

                ULTIMA_LISTA_QUE_HIZO_CLIC = Mid(CODIGO_REPORTE, 1, 1)
                GoTo SALIR
            End If
            If (CODIGO_REPORTE = "S0005" Or CODIGO_REPORTE = "S0006") Then                      'S0005 - Subsidios por Niveles Organizacionales;  S0006 - Detalles de Subsidios por Niveles Organizacionales
                '--------------- NIVELES ORGANIZACIONALES
                Me.ComboBox1.Enabled = True        'NIVEL 1
                Me.ComboBox2.Enabled = True        'NIVEL 2
                Me.CheckBox1.Enabled = True        'CHECK N2
                Me.ComboBox3.Enabled = True        'NIVEL 3
                Me.CheckBox2.Enabled = True        'CHECK N3
                Me.ComboBox4.Enabled = True        'NIVEL 4
                Me.CheckBox3.Enabled = True        'CHECK N4
                Me.ComboBox5.Enabled = True        'NIVEL 5
                Me.CheckBox4.Enabled = True        'CHECK N5
                Me.ComboBox6.Enabled = True        'NIVEL 6
                Me.CheckBox5.Enabled = True        'CHECK N6
                Me.ComboBox7.Enabled = True        'NIVEL 7
                Me.CheckBox6.Enabled = True        'CHECK N7
                Me.ComboBox8.Enabled = True        'NIVEL 8
                Me.CheckBox7.Enabled = True        'CHECK N8
                Me.Button1.Enabled = True          'BOTON NIVELES ORGANIZACIONALES
                '--------------- TIPO DE ESTRUCTURAS
                Me.RadioButton1.Enabled = True     'ESTRUCTURA MIXTA
                Me.RadioButton2.Enabled = True     'ESTRUCTURA MILITAR
                Me.RadioButton3.Enabled = True     'ESTRUCTURA PAME
                Me.RadioButton4.Enabled = True     'JEFE
                Me.Button2.Enabled = True          'BOTON TIPO DE ESTRUCTURAS
                '--------------- GRADOS, CARGOS, OTROS
                Me.ComboBox9.Enabled = True        'GRADO REAL
                Me.Button3.Enabled = True          'BOTON GR
                Me.ComboBox10.Enabled = True       'GRADO PLANTILLA
                Me.Button4.Enabled = True          'BOTON GP
                Me.ComboBox11.Enabled = True       'SITUACION DEL CARGO
                Me.Button5.Enabled = True          'BOTON SITUACION CARGO
                Me.ComboBox12.Enabled = True       'CARGO
                Me.Button6.Enabled = True          'BOTON CARGO
                Me.ComboBox13.Enabled = True       'ESTABLECIMIENTO
                Me.Button7.Enabled = True          'BOTON ESTABLECIMIENTO
                '--------------- FECHAS
                'Me.RadioButton5.Enabled = True    'FECHA DE INGRESO PAME
                'Me.RadioButton5.Checked = True
                'Me.RadioButton6.Enabled = True    'FECHA DE INGRESO EN
                'Me.RadioButton7.Enabled = True    'FECHA DE NACIMIENTO
                'Me.RadioButton7.Checked = True
                'Me.RadioButton8.Enabled = True    'FECHA DE SITUACION
                'Me.RadioButton9.Enabled = True    'FECHA DE MARCA
                'Me.RadioButton9.Checked = True
                'Me.RadioButton10.Enabled = True   'FECHA DE PROCESO DEL ASPIRANTE
                'Me.RadioButton11.Enabled = True   'FECHA DE APLICACION DE MOVIMIENTOS
                'Me.RadioButton15.Enabled = True   'FECHA DE EVENTUALIDADES
                Me.RadioButton16.Enabled = True    'FECHA DE SUBSIDIOS
                Me.RadioButton16.Checked = True
                Me.DateTimePicker1.Enabled = True  'DESDE
                Me.DateTimePicker2.Enabled = True  'HASTA
                Me.Button11.Enabled = True         'BOTON FECHAS
                '--------------- OTROS FILTROS
                'Me.ComboBox19.Enabled = True      'FUNCIONALIDAD DE EMPLEADO
                'Me.Button14.Enabled = True        'BOTON FUNCIONALIDAD DE EMPLEADO
                'Me.ComboBox20.Enabled = True      'CAUSA REAL DE EVENTUALIDAD
                'Me.Button15.Enabled = True        'BOTON CAUSA REAL DE EVENTUALIDAD
                'Me.ComboBox21.Enabled = True      'TIPO DE EVENTUALIDAD
                'Me.Button16.Enabled = True        'BOTON TIPO DE EVENTUALIDAD
                'Me.ComboBox22.Enabled = True      'ESTUDIOS
                'Me.Button17.Enabled = True        'BOTON ESTUDIOS
                'Me.ComboBox23.Enabled = True      'NIVEL DE COMPETENCIA
                'Me.Button18.Enabled = True        'BOTON NIVEL DE COMPETENCIA
                'Me.ComboBox24.Enabled = True      'RELOJ DE MARCACION
                'Me.Button19.Enabled = True        'BOTON RELOJ DE MARCACION
                'Me.ComboBox25.Enabled = True      'TIPO DE MARCA
                'Me.Button20.Enabled = True        'BOTON TIPO DE MARCA
                'Me.ComboBox26.Enabled = True      'CLASIFICACION DEL EMPLEADO
                'Me.Button21.Enabled = True        'BOTON CLASIFICACION
                'Me.TextBox1.Enabled = True        'CODIGO DEL EMPLEADO
                'Me.Button22.Enabled = True        'BOTON CODIGO EMPLEADO
                'Me.RadioButton12.Enabled = True   'EDAD
                'Me.RadioButton13.Enabled = True   'ANTIGUEDAD PAME
                'Me.RadioButton14.Enabled = True   'ANTIGUEDAD EN
                'Me.TextBox2.Enabled = True        'ANTIGUEDADES A
                'Me.TextBox3.Enabled = True        'ANTIGUEDADES B
                'Me.Button23.Enabled = True        'BOTON ANTIGUEDADES
                'Me.TextBox4.Enabled = True        'CONTRATOS
                'Me.Button24.Enabled = True        'BOTON CONTRATOS
                '--------------- CONDECORACIONES
                'Me.ComboBox27.Enabled = True      'CLASIFICACION
                'Me.Button25.Enabled = True        'BOTON CLASIFICACION
                'Me.ComboBox28.Enabled = True      'TIPO
                'Me.Button26.Enabled = True        'BOTON TIPO
                '--------------- SITUACIONES
                'Me.ComboBox29.Enabled = True      'TIPO SITUACION
                'Me.Button27.Enabled = True        'BOTON TIPO SITUACION
                'Me.ComboBox30.Enabled = True      'SITUACION
                'Me.Button28.Enabled = True        'BOTON SITUACION
                'Me.DateTimePicker3.Enabled = True 'DESDE
                'Me.DateTimePicker4.Enabled = True 'HASTA
                'Me.Button29.Enabled = True        'BOTON FECHAS

                ULTIMA_LISTA_QUE_HIZO_CLIC = Mid(CODIGO_REPORTE, 1, 1)
                GoTo SALIR
            End If
            If (CODIGO_REPORTE = "S0007") Then                      'S0007 - Situaciones del Personal por Rango de Fechas
                '--------------- NIVELES ORGANIZACIONALES
                'Me.ComboBox1.Enabled = True        'NIVEL 1
                'Me.ComboBox2.Enabled = True        'NIVEL 2
                'Me.CheckBox1.Enabled = True        'CHECK N2
                'Me.ComboBox3.Enabled = True        'NIVEL 3
                'Me.CheckBox2.Enabled = True        'CHECK N3
                'Me.ComboBox4.Enabled = True        'NIVEL 4
                'Me.CheckBox3.Enabled = True        'CHECK N4
                'Me.ComboBox5.Enabled = True        'NIVEL 5
                'Me.CheckBox4.Enabled = True        'CHECK N5
                'Me.ComboBox6.Enabled = True        'NIVEL 6
                'Me.CheckBox5.Enabled = True        'CHECK N6
                'Me.ComboBox7.Enabled = True        'NIVEL 7
                'Me.CheckBox6.Enabled = True        'CHECK N7
                'Me.ComboBox8.Enabled = True        'NIVEL 8
                'Me.CheckBox7.Enabled = True        'CHECK N8
                'Me.Button1.Enabled = True          'BOTON NIVELES ORGANIZACIONALES
                '--------------- TIPO DE ESTRUCTURAS
                'Me.RadioButton1.Enabled = True     'ESTRUCTURA MIXTA
                'Me.RadioButton2.Enabled = True     'ESTRUCTURA MILITAR
                'Me.RadioButton3.Enabled = True     'ESTRUCTURA PAME
                'Me.RadioButton4.Enabled = True     'JEFE
                'Me.Button2.Enabled = True          'BOTON TIPO DE ESTRUCTURAS
                '--------------- GRADOS, CARGOS, OTROS
                'Me.ComboBox9.Enabled = True        'GRADO REAL
                'Me.Button3.Enabled = True          'BOTON GR
                'Me.ComboBox10.Enabled = True       'GRADO PLANTILLA
                'Me.Button4.Enabled = True          'BOTON GP
                'Me.ComboBox11.Enabled = True       'SITUACION DEL CARGO
                'Me.Button5.Enabled = True          'BOTON SITUACION CARGO
                'Me.ComboBox12.Enabled = True       'CARGO
                'Me.Button6.Enabled = True          'BOTON CARGO
                'Me.ComboBox13.Enabled = True       'ESTABLECIMIENTO
                'Me.Button7.Enabled = True          'BOTON ESTABLECIMIENTO
                '--------------- FECHAS
                'Me.RadioButton5.Enabled = True    'FECHA DE INGRESO PAME
                'Me.RadioButton5.Checked = True
                'Me.RadioButton6.Enabled = True    'FECHA DE INGRESO EN
                'Me.RadioButton7.Enabled = True    'FECHA DE NACIMIENTO
                'Me.RadioButton7.Checked = True
                'Me.RadioButton8.Enabled = True    'FECHA DE SITUACION
                'Me.RadioButton9.Enabled = True    'FECHA DE MARCA
                'Me.RadioButton9.Checked = True
                'Me.RadioButton10.Enabled = True   'FECHA DE PROCESO DEL ASPIRANTE
                'Me.RadioButton11.Enabled = True   'FECHA DE APLICACION DE MOVIMIENTOS
                'Me.RadioButton15.Enabled = True   'FECHA DE EVENTUALIDADES
                'Me.RadioButton16.Enabled = True   'FECHA DE SUBSIDIOS
                'Me.RadioButton16.Checked = True
                'Me.DateTimePicker1.Enabled = True 'DESDE
                'Me.DateTimePicker2.Enabled = True 'HASTA
                'Me.Button11.Enabled = True        'BOTON FECHAS
                '--------------- OTROS FILTROS
                'Me.ComboBox19.Enabled = True      'FUNCIONALIDAD DE EMPLEADO
                'Me.Button14.Enabled = True        'BOTON FUNCIONALIDAD DE EMPLEADO
                'Me.ComboBox20.Enabled = True      'CAUSA REAL DE EVENTUALIDAD
                'Me.Button15.Enabled = True        'BOTON CAUSA REAL DE EVENTUALIDAD
                'Me.ComboBox21.Enabled = True      'TIPO DE EVENTUALIDAD
                'Me.Button16.Enabled = True        'BOTON TIPO DE EVENTUALIDAD
                'Me.ComboBox22.Enabled = True      'ESTUDIOS
                'Me.Button17.Enabled = True        'BOTON ESTUDIOS
                'Me.ComboBox23.Enabled = True      'NIVEL DE COMPETENCIA
                'Me.Button18.Enabled = True        'BOTON NIVEL DE COMPETENCIA
                'Me.ComboBox24.Enabled = True      'RELOJ DE MARCACION
                'Me.Button19.Enabled = True        'BOTON RELOJ DE MARCACION
                'Me.ComboBox25.Enabled = True      'TIPO DE MARCA
                'Me.Button20.Enabled = True        'BOTON TIPO DE MARCA
                'Me.ComboBox26.Enabled = True      'CLASIFICACION DEL EMPLEADO
                'Me.Button21.Enabled = True        'BOTON CLASIFICACION
                Me.TextBox1.Enabled = True        'CODIGO DEL EMPLEADO
                Me.Button22.Enabled = True        'BOTON CODIGO EMPLEADO
                'Me.RadioButton12.Enabled = True   'EDAD
                'Me.RadioButton13.Enabled = True   'ANTIGUEDAD PAME
                'Me.RadioButton14.Enabled = True   'ANTIGUEDAD EN
                'Me.TextBox2.Enabled = True        'ANTIGUEDADES A
                'Me.TextBox3.Enabled = True        'ANTIGUEDADES B
                'Me.Button23.Enabled = True        'BOTON ANTIGUEDADES
                'Me.TextBox4.Enabled = True        'CONTRATOS
                'Me.Button24.Enabled = True        'BOTON CONTRATOS
                '--------------- CONDECORACIONES
                'Me.ComboBox27.Enabled = True      'CLASIFICACION
                'Me.Button25.Enabled = True        'BOTON CLASIFICACION
                'Me.ComboBox28.Enabled = True      'TIPO
                'Me.Button26.Enabled = True        'BOTON TIPO
                '--------------- SITUACIONES
                Me.ComboBox29.Enabled = True      'TIPO SITUACION
                Me.Button27.Enabled = True        'BOTON TIPO SITUACION
                Me.ComboBox30.Enabled = True      'SITUACION
                Me.Button28.Enabled = True        'BOTON SITUACION
                Me.DateTimePicker3.Enabled = True 'DESDE
                Me.DateTimePicker4.Enabled = True 'HASTA
                Me.Button29.Enabled = True        'BOTON FECHAS

                ULTIMA_LISTA_QUE_HIZO_CLIC = Mid(CODIGO_REPORTE, 1, 1)
                GoTo SALIR
            End If
            If (CODIGO_REPORTE = "S0008") Then                      'S0008 - Estado Vacacional por Rango de Fechas (Detalles)
                '--------------- NIVELES ORGANIZACIONALES
                Me.ComboBox1.Enabled = True        'NIVEL 1
                Me.ComboBox2.Enabled = True        'NIVEL 2
                Me.CheckBox1.Enabled = True        'CHECK N2
                Me.ComboBox3.Enabled = True        'NIVEL 3
                Me.CheckBox2.Enabled = True        'CHECK N3
                Me.ComboBox4.Enabled = True        'NIVEL 4
                Me.CheckBox3.Enabled = True        'CHECK N4
                Me.ComboBox5.Enabled = True        'NIVEL 5
                Me.CheckBox4.Enabled = True        'CHECK N5
                Me.ComboBox6.Enabled = True        'NIVEL 6
                Me.CheckBox5.Enabled = True        'CHECK N6
                Me.ComboBox7.Enabled = True        'NIVEL 7
                Me.CheckBox6.Enabled = True        'CHECK N7
                Me.ComboBox8.Enabled = True        'NIVEL 8
                Me.CheckBox7.Enabled = True        'CHECK N8
                Me.Button1.Enabled = True          'BOTON NIVELES ORGANIZACIONALES
                '--------------- TIPO DE ESTRUCTURAS
                Me.RadioButton1.Enabled = True     'ESTRUCTURA MIXTA
                Me.RadioButton2.Enabled = True     'ESTRUCTURA MILITAR
                Me.RadioButton3.Enabled = True     'ESTRUCTURA PAME
                Me.RadioButton4.Enabled = True     'JEFE
                Me.Button2.Enabled = True          'BOTON TIPO DE ESTRUCTURAS
                '--------------- GRADOS, CARGOS, OTROS
                Me.ComboBox9.Enabled = True        'GRADO REAL
                Me.Button3.Enabled = True          'BOTON GR
                Me.ComboBox10.Enabled = True       'GRADO PLANTILLA
                Me.Button4.Enabled = True          'BOTON GP
                Me.ComboBox11.Enabled = True       'SITUACION DEL CARGO
                Me.Button5.Enabled = True          'BOTON SITUACION CARGO
                Me.ComboBox12.Enabled = True       'CARGO
                Me.Button6.Enabled = True          'BOTON CARGO
                Me.ComboBox13.Enabled = True       'ESTABLECIMIENTO
                Me.Button7.Enabled = True          'BOTON ESTABLECIMIENTO
                '--------------- FECHAS
                'Me.RadioButton5.Enabled = True    'FECHA DE INGRESO PAME
                'Me.RadioButton5.Checked = True
                'Me.RadioButton6.Enabled = True    'FECHA DE INGRESO EN
                'Me.RadioButton7.Enabled = True    'FECHA DE NACIMIENTO
                'Me.RadioButton7.Checked = True
                'Me.RadioButton8.Enabled = True    'FECHA DE SITUACION
                'Me.RadioButton9.Enabled = True    'FECHA DE MARCA
                'Me.RadioButton9.Checked = True
                'Me.RadioButton10.Enabled = True   'FECHA DE PROCESO DEL ASPIRANTE
                'Me.RadioButton11.Enabled = True   'FECHA DE APLICACION DE MOVIMIENTOS
                'Me.RadioButton15.Enabled = True   'FECHA DE EVENTUALIDADES
                'Me.RadioButton16.Enabled = True   'FECHA DE SUBSIDIOS
                'Me.RadioButton16.Checked = True
                'Me.DateTimePicker1.Enabled = True 'DESDE
                'Me.DateTimePicker2.Enabled = True 'HASTA
                'Me.Button11.Enabled = True        'BOTON FECHAS
                '--------------- OTROS FILTROS
                'Me.ComboBox19.Enabled = True      'FUNCIONALIDAD DE EMPLEADO
                'Me.Button14.Enabled = True        'BOTON FUNCIONALIDAD DE EMPLEADO
                'Me.ComboBox20.Enabled = True      'CAUSA REAL DE EVENTUALIDAD
                'Me.Button15.Enabled = True        'BOTON CAUSA REAL DE EVENTUALIDAD
                'Me.ComboBox21.Enabled = True      'TIPO DE EVENTUALIDAD
                'Me.Button16.Enabled = True        'BOTON TIPO DE EVENTUALIDAD
                'Me.ComboBox22.Enabled = True      'ESTUDIOS
                'Me.Button17.Enabled = True        'BOTON ESTUDIOS
                'Me.ComboBox23.Enabled = True      'NIVEL DE COMPETENCIA
                'Me.Button18.Enabled = True        'BOTON NIVEL DE COMPETENCIA
                'Me.ComboBox24.Enabled = True      'RELOJ DE MARCACION
                'Me.Button19.Enabled = True        'BOTON RELOJ DE MARCACION
                'Me.ComboBox25.Enabled = True      'TIPO DE MARCA
                'Me.Button20.Enabled = True        'BOTON TIPO DE MARCA
                'Me.ComboBox26.Enabled = True      'CLASIFICACION DEL EMPLEADO
                'Me.Button21.Enabled = True        'BOTON CLASIFICACION
                'Me.TextBox1.Enabled = True        'CODIGO DEL EMPLEADO
                'Me.Button22.Enabled = True        'BOTON CODIGO EMPLEADO
                'Me.RadioButton12.Enabled = True   'EDAD
                'Me.RadioButton13.Enabled = True   'ANTIGUEDAD PAME
                'Me.RadioButton14.Enabled = True   'ANTIGUEDAD EN
                'Me.TextBox2.Enabled = True        'ANTIGUEDADES A
                'Me.TextBox3.Enabled = True        'ANTIGUEDADES B
                'Me.Button23.Enabled = True        'BOTON ANTIGUEDADES
                'Me.TextBox4.Enabled = True        'CONTRATOS
                'Me.Button24.Enabled = True        'BOTON CONTRATOS
                '--------------- CONDECORACIONES
                'Me.ComboBox27.Enabled = True      'CLASIFICACION
                'Me.Button25.Enabled = True        'BOTON CLASIFICACION
                'Me.ComboBox28.Enabled = True      'TIPO
                'Me.Button26.Enabled = True        'BOTON TIPO
                '--------------- SITUACIONES
                'Me.ComboBox29.Enabled = True      'TIPO SITUACION
                Me.ComboBox29.Text = "AUSENTE"
                'Me.Button27.Enabled = True        'BOTON TIPO SITUACION
                'Me.ComboBox30.Enabled = True      'SITUACION
                Me.ComboBox30.Text = "VACACIONES"
                'Me.Button28.Enabled = True        'BOTON SITUACION
                Me.DateTimePicker3.Enabled = True  'DESDE
                Me.DateTimePicker4.Enabled = True  'HASTA
                Me.Button29.Enabled = True         'BOTON FECHAS
                ULTIMA_LISTA_QUE_HIZO_CLIC = Mid(CODIGO_REPORTE, 1, 1)
                GoTo SALIR
            End If
            If (CODIGO_REPORTE = "S0009") Then                      'S0009 - Estado Vacacional por Rango de Fechas (Niveles Organizacionales)
                '--------------- NIVELES ORGANIZACIONALES
                Me.ComboBox1.Enabled = True        'NIVEL 1
                Me.ComboBox2.Enabled = True        'NIVEL 2
                Me.CheckBox1.Enabled = True        'CHECK N2
                Me.ComboBox3.Enabled = True        'NIVEL 3
                Me.CheckBox2.Enabled = True        'CHECK N3
                Me.ComboBox4.Enabled = True        'NIVEL 4
                Me.CheckBox3.Enabled = True        'CHECK N4
                Me.ComboBox5.Enabled = True        'NIVEL 5
                Me.CheckBox4.Enabled = True        'CHECK N5
                Me.ComboBox6.Enabled = True        'NIVEL 6
                Me.CheckBox5.Enabled = True        'CHECK N6
                Me.ComboBox7.Enabled = True        'NIVEL 7
                Me.CheckBox6.Enabled = True        'CHECK N7
                Me.ComboBox8.Enabled = True        'NIVEL 8
                Me.CheckBox7.Enabled = True        'CHECK N8
                Me.Button1.Enabled = True          'BOTON NIVELES ORGANIZACIONALES
                '--------------- TIPO DE ESTRUCTURAS
                Me.RadioButton1.Enabled = True     'ESTRUCTURA MIXTA
                Me.RadioButton2.Enabled = True     'ESTRUCTURA MILITAR
                Me.RadioButton3.Enabled = True     'ESTRUCTURA PAME
                Me.RadioButton4.Enabled = True     'JEFE
                Me.Button2.Enabled = True          'BOTON TIPO DE ESTRUCTURAS
                '--------------- GRADOS, CARGOS, OTROS
                Me.ComboBox9.Enabled = True        'GRADO REAL
                Me.Button3.Enabled = True          'BOTON GR
                Me.ComboBox10.Enabled = True       'GRADO PLANTILLA
                Me.Button4.Enabled = True          'BOTON GP
                Me.ComboBox11.Enabled = True       'SITUACION DEL CARGO
                Me.Button5.Enabled = True          'BOTON SITUACION CARGO
                Me.ComboBox12.Enabled = True       'CARGO
                Me.Button6.Enabled = True          'BOTON CARGO
                Me.ComboBox13.Enabled = True       'ESTABLECIMIENTO
                Me.Button7.Enabled = True          'BOTON ESTABLECIMIENTO
                '--------------- FECHAS
                'Me.RadioButton5.Enabled = True    'FECHA DE INGRESO PAME
                'Me.RadioButton5.Checked = True
                'Me.RadioButton6.Enabled = True    'FECHA DE INGRESO EN
                'Me.RadioButton7.Enabled = True    'FECHA DE NACIMIENTO
                'Me.RadioButton7.Checked = True
                'Me.RadioButton8.Enabled = True    'FECHA DE SITUACION
                'Me.RadioButton9.Enabled = True    'FECHA DE MARCA
                'Me.RadioButton9.Checked = True
                'Me.RadioButton10.Enabled = True   'FECHA DE PROCESO DEL ASPIRANTE
                'Me.RadioButton11.Enabled = True   'FECHA DE APLICACION DE MOVIMIENTOS
                'Me.RadioButton15.Enabled = True   'FECHA DE EVENTUALIDADES
                'Me.RadioButton16.Enabled = True   'FECHA DE SUBSIDIOS
                'Me.RadioButton16.Checked = True
                'Me.DateTimePicker1.Enabled = True 'DESDE
                'Me.DateTimePicker2.Enabled = True 'HASTA
                'Me.Button11.Enabled = True        'BOTON FECHAS
                '--------------- OTROS FILTROS
                'Me.ComboBox19.Enabled = True      'FUNCIONALIDAD DE EMPLEADO
                'Me.Button14.Enabled = True        'BOTON FUNCIONALIDAD DE EMPLEADO
                'Me.ComboBox20.Enabled = True      'CAUSA REAL DE EVENTUALIDAD
                'Me.Button15.Enabled = True        'BOTON CAUSA REAL DE EVENTUALIDAD
                'Me.ComboBox21.Enabled = True      'TIPO DE EVENTUALIDAD
                'Me.Button16.Enabled = True        'BOTON TIPO DE EVENTUALIDAD
                'Me.ComboBox22.Enabled = True      'ESTUDIOS
                'Me.Button17.Enabled = True        'BOTON ESTUDIOS
                'Me.ComboBox23.Enabled = True      'NIVEL DE COMPETENCIA
                'Me.Button18.Enabled = True        'BOTON NIVEL DE COMPETENCIA
                'Me.ComboBox24.Enabled = True      'RELOJ DE MARCACION
                'Me.Button19.Enabled = True        'BOTON RELOJ DE MARCACION
                'Me.ComboBox25.Enabled = True      'TIPO DE MARCA
                'Me.Button20.Enabled = True        'BOTON TIPO DE MARCA
                'Me.ComboBox26.Enabled = True      'CLASIFICACION DEL EMPLEADO
                'Me.Button21.Enabled = True        'BOTON CLASIFICACION
                'Me.TextBox1.Enabled = True        'CODIGO DEL EMPLEADO
                'Me.Button22.Enabled = True        'BOTON CODIGO EMPLEADO
                'Me.RadioButton12.Enabled = True   'EDAD
                'Me.RadioButton13.Enabled = True   'ANTIGUEDAD PAME
                'Me.RadioButton14.Enabled = True   'ANTIGUEDAD EN
                'Me.TextBox2.Enabled = True        'ANTIGUEDADES A
                'Me.TextBox3.Enabled = True        'ANTIGUEDADES B
                'Me.Button23.Enabled = True        'BOTON ANTIGUEDADES
                'Me.TextBox4.Enabled = True        'CONTRATOS
                'Me.Button24.Enabled = True        'BOTON CONTRATOS
                '--------------- CONDECORACIONES
                'Me.ComboBox27.Enabled = True      'CLASIFICACION
                'Me.Button25.Enabled = True        'BOTON CLASIFICACION
                'Me.ComboBox28.Enabled = True      'TIPO
                'Me.Button26.Enabled = True        'BOTON TIPO
                '--------------- SITUACIONES
                'Me.ComboBox29.Enabled = True      'TIPO SITUACION
                Me.ComboBox29.Text = "AUSENTE"
                'Me.Button27.Enabled = True        'BOTON TIPO SITUACION
                'Me.ComboBox30.Enabled = True      'SITUACION
                Me.ComboBox30.Text = "VACACIONES"
                'Me.Button28.Enabled = True        'BOTON SITUACION
                Me.DateTimePicker3.Enabled = True  'DESDE
                Me.DateTimePicker4.Enabled = True  'HASTA
                Me.Button29.Enabled = True         'BOTON FECHAS
                ULTIMA_LISTA_QUE_HIZO_CLIC = Mid(CODIGO_REPORTE, 1, 1)
                GoTo SALIR
            End If
            If (CODIGO_REPORTE = "S0010") Then                      'S0010 - Indicadores de < Cantidad de Empleados por Tipo de Situacion A >
                '--------------- NIVELES ORGANIZACIONALES
                Me.ComboBox1.Enabled = True        'NIVEL 1
                Me.ComboBox2.Enabled = True        'NIVEL 2
                Me.CheckBox1.Enabled = True        'CHECK N2
                Me.ComboBox3.Enabled = True        'NIVEL 3
                Me.CheckBox2.Enabled = True        'CHECK N3
                Me.ComboBox4.Enabled = True        'NIVEL 4
                Me.CheckBox3.Enabled = True        'CHECK N4
                Me.ComboBox5.Enabled = True        'NIVEL 5
                Me.CheckBox4.Enabled = True        'CHECK N5
                Me.ComboBox6.Enabled = True        'NIVEL 6
                Me.CheckBox5.Enabled = True        'CHECK N6
                Me.ComboBox7.Enabled = True        'NIVEL 7
                Me.CheckBox6.Enabled = True        'CHECK N7
                Me.ComboBox8.Enabled = True        'NIVEL 8
                Me.CheckBox7.Enabled = True        'CHECK N8
                Me.Button1.Enabled = True          'BOTON NIVELES ORGANIZACIONALES
                '--------------- TIPO DE ESTRUCTURAS
                Me.RadioButton1.Enabled = True     'ESTRUCTURA MIXTA
                Me.RadioButton2.Enabled = True     'ESTRUCTURA MILITAR
                Me.RadioButton3.Enabled = True     'ESTRUCTURA PAME
                Me.RadioButton4.Enabled = True     'JEFE
                Me.Button2.Enabled = True          'BOTON TIPO DE ESTRUCTURAS
                '--------------- GRADOS, CARGOS, OTROS
                Me.ComboBox9.Enabled = True        'GRADO REAL
                Me.Button3.Enabled = True          'BOTON GR
                Me.ComboBox10.Enabled = True       'GRADO PLANTILLA
                Me.Button4.Enabled = True          'BOTON GP
                Me.ComboBox11.Enabled = True       'SITUACION DEL CARGO
                Me.Button5.Enabled = True          'BOTON SITUACION CARGO
                Me.ComboBox12.Enabled = True       'CARGO
                Me.Button6.Enabled = True          'BOTON CARGO
                Me.ComboBox13.Enabled = True       'ESTABLECIMIENTO
                Me.Button7.Enabled = True          'BOTON ESTABLECIMIENTO
                '--------------- FECHAS
                'Me.RadioButton5.Enabled = True    'FECHA DE INGRESO PAME
                'Me.RadioButton5.Checked = True
                'Me.RadioButton6.Enabled = True    'FECHA DE INGRESO EN
                'Me.RadioButton7.Enabled = True    'FECHA DE NACIMIENTO
                'Me.RadioButton7.Checked = True
                'Me.RadioButton8.Enabled = True    'FECHA DE SITUACION
                'Me.RadioButton9.Enabled = True    'FECHA DE MARCA
                'Me.RadioButton9.Checked = True
                'Me.RadioButton10.Enabled = True   'FECHA DE PROCESO DEL ASPIRANTE
                'Me.RadioButton11.Enabled = True   'FECHA DE APLICACION DE MOVIMIENTOS
                'Me.RadioButton15.Enabled = True   'FECHA DE EVENTUALIDADES
                'Me.RadioButton16.Enabled = True   'FECHA DE SUBSIDIOS
                'Me.RadioButton16.Checked = True
                'Me.DateTimePicker1.Enabled = True 'DESDE
                'Me.DateTimePicker2.Enabled = True 'HASTA
                'Me.Button11.Enabled = True        'BOTON FECHAS
                '--------------- OTROS FILTROS
                'Me.ComboBox19.Enabled = True      'FUNCIONALIDAD DE EMPLEADO
                'Me.Button14.Enabled = True        'BOTON FUNCIONALIDAD DE EMPLEADO
                'Me.ComboBox20.Enabled = True      'CAUSA REAL DE EVENTUALIDAD
                'Me.Button15.Enabled = True        'BOTON CAUSA REAL DE EVENTUALIDAD
                'Me.ComboBox21.Enabled = True      'TIPO DE EVENTUALIDAD
                'Me.Button16.Enabled = True        'BOTON TIPO DE EVENTUALIDAD
                'Me.ComboBox22.Enabled = True      'ESTUDIOS
                'Me.Button17.Enabled = True        'BOTON ESTUDIOS
                'Me.ComboBox23.Enabled = True      'NIVEL DE COMPETENCIA
                'Me.Button18.Enabled = True        'BOTON NIVEL DE COMPETENCIA
                'Me.ComboBox24.Enabled = True      'RELOJ DE MARCACION
                'Me.Button19.Enabled = True        'BOTON RELOJ DE MARCACION
                'Me.ComboBox25.Enabled = True      'TIPO DE MARCA
                'Me.Button20.Enabled = True        'BOTON TIPO DE MARCA
                'Me.ComboBox26.Enabled = True      'CLASIFICACION DEL EMPLEADO
                'Me.Button21.Enabled = True        'BOTON CLASIFICACION
                'Me.TextBox1.Enabled = True        'CODIGO DEL EMPLEADO
                'Me.Button22.Enabled = True        'BOTON CODIGO EMPLEADO
                'Me.RadioButton12.Enabled = True   'EDAD
                'Me.RadioButton13.Enabled = True   'ANTIGUEDAD PAME
                'Me.RadioButton14.Enabled = True   'ANTIGUEDAD EN
                'Me.TextBox2.Enabled = True        'ANTIGUEDADES A
                'Me.TextBox3.Enabled = True        'ANTIGUEDADES B
                'Me.Button23.Enabled = True        'BOTON ANTIGUEDADES
                'Me.TextBox4.Enabled = True        'CONTRATOS
                'Me.Button24.Enabled = True        'BOTON CONTRATOS
                '--------------- CONDECORACIONES
                'Me.ComboBox27.Enabled = True      'CLASIFICACION
                'Me.Button25.Enabled = True        'BOTON CLASIFICACION
                'Me.ComboBox28.Enabled = True      'TIPO
                'Me.Button26.Enabled = True        'BOTON TIPO
                '--------------- SITUACIONES
                Me.ComboBox29.Enabled = True      'TIPO SITUACION
                Me.ComboBox29.Text = "AUSENTE"
                Me.Button27.Enabled = True        'BOTON TIPO SITUACION
                Me.ComboBox30.Enabled = True      'SITUACION
                Me.ComboBox30.Text = "VACACIONES"
                Me.Button28.Enabled = True        'BOTON SITUACION
                Me.DateTimePicker3.Enabled = True  'DESDE
                Me.DateTimePicker4.Enabled = True  'HASTA
                Me.Button29.Enabled = True         'BOTON FECHAS
                ULTIMA_LISTA_QUE_HIZO_CLIC = Mid(CODIGO_REPORTE, 1, 1)
                GoTo SALIR
            End If
            If (CODIGO_REPORTE = "S0011") Then                      'S0011 - Indicadores de < Cantidad de Empleados por Tipo de Situacion B >
                '--------------- NIVELES ORGANIZACIONALES
                Me.ComboBox1.Enabled = True        'NIVEL 1
                Me.ComboBox2.Enabled = True        'NIVEL 2
                Me.CheckBox1.Enabled = True        'CHECK N2
                Me.ComboBox3.Enabled = True        'NIVEL 3
                Me.CheckBox2.Enabled = True        'CHECK N3
                Me.ComboBox4.Enabled = True        'NIVEL 4
                Me.CheckBox3.Enabled = True        'CHECK N4
                Me.ComboBox5.Enabled = True        'NIVEL 5
                Me.CheckBox4.Enabled = True        'CHECK N5
                Me.ComboBox6.Enabled = True        'NIVEL 6
                Me.CheckBox5.Enabled = True        'CHECK N6
                Me.ComboBox7.Enabled = True        'NIVEL 7
                Me.CheckBox6.Enabled = True        'CHECK N7
                Me.ComboBox8.Enabled = True        'NIVEL 8
                Me.CheckBox7.Enabled = True        'CHECK N8
                Me.Button1.Enabled = True          'BOTON NIVELES ORGANIZACIONALES
                '--------------- TIPO DE ESTRUCTURAS
                Me.RadioButton1.Enabled = True     'ESTRUCTURA MIXTA
                Me.RadioButton2.Enabled = True     'ESTRUCTURA MILITAR
                Me.RadioButton3.Enabled = True     'ESTRUCTURA PAME
                Me.RadioButton4.Enabled = True     'JEFE
                Me.Button2.Enabled = True          'BOTON TIPO DE ESTRUCTURAS
                '--------------- GRADOS, CARGOS, OTROS
                Me.ComboBox9.Enabled = True        'GRADO REAL
                Me.Button3.Enabled = True          'BOTON GR
                Me.ComboBox10.Enabled = True       'GRADO PLANTILLA
                Me.Button4.Enabled = True          'BOTON GP
                Me.ComboBox11.Enabled = True       'SITUACION DEL CARGO
                Me.Button5.Enabled = True          'BOTON SITUACION CARGO
                Me.ComboBox12.Enabled = True       'CARGO
                Me.Button6.Enabled = True          'BOTON CARGO
                Me.ComboBox13.Enabled = True       'ESTABLECIMIENTO
                Me.Button7.Enabled = True          'BOTON ESTABLECIMIENTO
                '--------------- FECHAS
                'Me.RadioButton5.Enabled = True    'FECHA DE INGRESO PAME
                'Me.RadioButton5.Checked = True
                'Me.RadioButton6.Enabled = True    'FECHA DE INGRESO EN
                'Me.RadioButton7.Enabled = True    'FECHA DE NACIMIENTO
                'Me.RadioButton7.Checked = True
                'Me.RadioButton8.Enabled = True    'FECHA DE SITUACION
                'Me.RadioButton9.Enabled = True    'FECHA DE MARCA
                'Me.RadioButton9.Checked = True
                'Me.RadioButton10.Enabled = True   'FECHA DE PROCESO DEL ASPIRANTE
                'Me.RadioButton11.Enabled = True   'FECHA DE APLICACION DE MOVIMIENTOS
                'Me.RadioButton15.Enabled = True   'FECHA DE EVENTUALIDADES
                'Me.RadioButton16.Enabled = True   'FECHA DE SUBSIDIOS
                'Me.RadioButton16.Checked = True
                'Me.DateTimePicker1.Enabled = True 'DESDE
                'Me.DateTimePicker2.Enabled = True 'HASTA
                'Me.Button11.Enabled = True        'BOTON FECHAS
                '--------------- OTROS FILTROS
                'Me.ComboBox19.Enabled = True      'FUNCIONALIDAD DE EMPLEADO
                'Me.Button14.Enabled = True        'BOTON FUNCIONALIDAD DE EMPLEADO
                'Me.ComboBox20.Enabled = True      'CAUSA REAL DE EVENTUALIDAD
                'Me.Button15.Enabled = True        'BOTON CAUSA REAL DE EVENTUALIDAD
                'Me.ComboBox21.Enabled = True      'TIPO DE EVENTUALIDAD
                'Me.Button16.Enabled = True        'BOTON TIPO DE EVENTUALIDAD
                'Me.ComboBox22.Enabled = True      'ESTUDIOS
                'Me.Button17.Enabled = True        'BOTON ESTUDIOS
                'Me.ComboBox23.Enabled = True      'NIVEL DE COMPETENCIA
                'Me.Button18.Enabled = True        'BOTON NIVEL DE COMPETENCIA
                'Me.ComboBox24.Enabled = True      'RELOJ DE MARCACION
                'Me.Button19.Enabled = True        'BOTON RELOJ DE MARCACION
                'Me.ComboBox25.Enabled = True      'TIPO DE MARCA
                'Me.Button20.Enabled = True        'BOTON TIPO DE MARCA
                'Me.ComboBox26.Enabled = True      'CLASIFICACION DEL EMPLEADO
                'Me.Button21.Enabled = True        'BOTON CLASIFICACION
                'Me.TextBox1.Enabled = True        'CODIGO DEL EMPLEADO
                'Me.Button22.Enabled = True        'BOTON CODIGO EMPLEADO
                'Me.RadioButton12.Enabled = True   'EDAD
                'Me.RadioButton13.Enabled = True   'ANTIGUEDAD PAME
                'Me.RadioButton14.Enabled = True   'ANTIGUEDAD EN
                'Me.TextBox2.Enabled = True        'ANTIGUEDADES A
                'Me.TextBox3.Enabled = True        'ANTIGUEDADES B
                'Me.Button23.Enabled = True        'BOTON ANTIGUEDADES
                'Me.TextBox4.Enabled = True        'CONTRATOS
                'Me.Button24.Enabled = True        'BOTON CONTRATOS
                '--------------- CONDECORACIONES
                'Me.ComboBox27.Enabled = True      'CLASIFICACION
                'Me.Button25.Enabled = True        'BOTON CLASIFICACION
                'Me.ComboBox28.Enabled = True      'TIPO
                'Me.Button26.Enabled = True        'BOTON TIPO
                '--------------- SITUACIONES
                Me.ComboBox29.Enabled = True      'TIPO SITUACION
                Me.ComboBox29.Text = "AUSENTE"
                Me.Button27.Enabled = True        'BOTON TIPO SITUACION
                Me.ComboBox30.Enabled = True      'SITUACION
                Me.ComboBox30.Text = "VACACIONES"
                Me.Button28.Enabled = True        'BOTON SITUACION
                Me.DateTimePicker3.Enabled = True  'DESDE
                Me.DateTimePicker4.Enabled = True  'HASTA
                Me.Button29.Enabled = True         'BOTON FECHAS
                ULTIMA_LISTA_QUE_HIZO_CLIC = Mid(CODIGO_REPORTE, 1, 1)
                GoTo SALIR
            End If
            If (CODIGO_REPORTE = "S0012") Then                      'S0012 - Indicadores de < Cantidad de Dias por Empleado y Tipo de Situacion >
                '--------------- NIVELES ORGANIZACIONALES
                Me.ComboBox1.Enabled = True        'NIVEL 1
                Me.ComboBox2.Enabled = True        'NIVEL 2
                Me.CheckBox1.Enabled = True        'CHECK N2
                Me.ComboBox3.Enabled = True        'NIVEL 3
                Me.CheckBox2.Enabled = True        'CHECK N3
                Me.ComboBox4.Enabled = True        'NIVEL 4
                Me.CheckBox3.Enabled = True        'CHECK N4
                Me.ComboBox5.Enabled = True        'NIVEL 5
                Me.CheckBox4.Enabled = True        'CHECK N5
                Me.ComboBox6.Enabled = True        'NIVEL 6
                Me.CheckBox5.Enabled = True        'CHECK N6
                Me.ComboBox7.Enabled = True        'NIVEL 7
                Me.CheckBox6.Enabled = True        'CHECK N7
                Me.ComboBox8.Enabled = True        'NIVEL 8
                Me.CheckBox7.Enabled = True        'CHECK N8
                Me.Button1.Enabled = True          'BOTON NIVELES ORGANIZACIONALES
                '--------------- TIPO DE ESTRUCTURAS
                Me.RadioButton1.Enabled = True     'ESTRUCTURA MIXTA
                Me.RadioButton2.Enabled = True     'ESTRUCTURA MILITAR
                Me.RadioButton3.Enabled = True     'ESTRUCTURA PAME
                Me.RadioButton4.Enabled = True     'JEFE
                Me.Button2.Enabled = True          'BOTON TIPO DE ESTRUCTURAS
                '--------------- GRADOS, CARGOS, OTROS
                Me.ComboBox9.Enabled = True        'GRADO REAL
                Me.Button3.Enabled = True          'BOTON GR
                Me.ComboBox10.Enabled = True       'GRADO PLANTILLA
                Me.Button4.Enabled = True          'BOTON GP
                Me.ComboBox11.Enabled = True       'SITUACION DEL CARGO
                Me.Button5.Enabled = True          'BOTON SITUACION CARGO
                Me.ComboBox12.Enabled = True       'CARGO
                Me.Button6.Enabled = True          'BOTON CARGO
                Me.ComboBox13.Enabled = True       'ESTABLECIMIENTO
                Me.Button7.Enabled = True          'BOTON ESTABLECIMIENTO
                '--------------- FECHAS
                'Me.RadioButton5.Enabled = True    'FECHA DE INGRESO PAME
                'Me.RadioButton5.Checked = True
                'Me.RadioButton6.Enabled = True    'FECHA DE INGRESO EN
                'Me.RadioButton7.Enabled = True    'FECHA DE NACIMIENTO
                'Me.RadioButton7.Checked = True
                'Me.RadioButton8.Enabled = True    'FECHA DE SITUACION
                'Me.RadioButton9.Enabled = True    'FECHA DE MARCA
                'Me.RadioButton9.Checked = True
                'Me.RadioButton10.Enabled = True   'FECHA DE PROCESO DEL ASPIRANTE
                'Me.RadioButton11.Enabled = True   'FECHA DE APLICACION DE MOVIMIENTOS
                'Me.RadioButton15.Enabled = True   'FECHA DE EVENTUALIDADES
                'Me.RadioButton16.Enabled = True   'FECHA DE SUBSIDIOS
                'Me.RadioButton16.Checked = True
                'Me.DateTimePicker1.Enabled = True 'DESDE
                'Me.DateTimePicker2.Enabled = True 'HASTA
                'Me.Button11.Enabled = True        'BOTON FECHAS
                '--------------- OTROS FILTROS
                'Me.ComboBox19.Enabled = True      'FUNCIONALIDAD DE EMPLEADO
                'Me.Button14.Enabled = True        'BOTON FUNCIONALIDAD DE EMPLEADO
                'Me.ComboBox20.Enabled = True      'CAUSA REAL DE EVENTUALIDAD
                'Me.Button15.Enabled = True        'BOTON CAUSA REAL DE EVENTUALIDAD
                'Me.ComboBox21.Enabled = True      'TIPO DE EVENTUALIDAD
                'Me.Button16.Enabled = True        'BOTON TIPO DE EVENTUALIDAD
                'Me.ComboBox22.Enabled = True      'ESTUDIOS
                'Me.Button17.Enabled = True        'BOTON ESTUDIOS
                'Me.ComboBox23.Enabled = True      'NIVEL DE COMPETENCIA
                'Me.Button18.Enabled = True        'BOTON NIVEL DE COMPETENCIA
                'Me.ComboBox24.Enabled = True      'RELOJ DE MARCACION
                'Me.Button19.Enabled = True        'BOTON RELOJ DE MARCACION
                'Me.ComboBox25.Enabled = True      'TIPO DE MARCA
                'Me.Button20.Enabled = True        'BOTON TIPO DE MARCA
                'Me.ComboBox26.Enabled = True      'CLASIFICACION DEL EMPLEADO
                'Me.Button21.Enabled = True        'BOTON CLASIFICACION
                'Me.TextBox1.Enabled = True        'CODIGO DEL EMPLEADO
                'Me.Button22.Enabled = True        'BOTON CODIGO EMPLEADO
                'Me.RadioButton12.Enabled = True   'EDAD
                'Me.RadioButton13.Enabled = True   'ANTIGUEDAD PAME
                'Me.RadioButton14.Enabled = True   'ANTIGUEDAD EN
                'Me.TextBox2.Enabled = True        'ANTIGUEDADES A
                'Me.TextBox3.Enabled = True        'ANTIGUEDADES B
                'Me.Button23.Enabled = True        'BOTON ANTIGUEDADES
                'Me.TextBox4.Enabled = True        'CONTRATOS
                'Me.Button24.Enabled = True        'BOTON CONTRATOS
                '--------------- CONDECORACIONES
                'Me.ComboBox27.Enabled = True      'CLASIFICACION
                'Me.Button25.Enabled = True        'BOTON CLASIFICACION
                'Me.ComboBox28.Enabled = True      'TIPO
                'Me.Button26.Enabled = True        'BOTON TIPO
                '--------------- SITUACIONES
                Me.ComboBox29.Enabled = True      'TIPO SITUACION
                Me.ComboBox29.Text = "AUSENTE"
                Me.Button27.Enabled = True        'BOTON TIPO SITUACION
                Me.ComboBox30.Enabled = True      'SITUACION
                Me.ComboBox30.Text = "VACACIONES"
                Me.Button28.Enabled = True        'BOTON SITUACION
                Me.DateTimePicker3.Enabled = True  'DESDE
                Me.DateTimePicker4.Enabled = True  'HASTA
                Me.Button29.Enabled = True         'BOTON FECHAS
                ULTIMA_LISTA_QUE_HIZO_CLIC = Mid(CODIGO_REPORTE, 1, 1)
                GoTo SALIR
            End If
            If (CODIGO_REPORTE = "S0013") Then                      'S0013 - Indicadores de < Lista de Empleados por Tipo de Situacion >
                '--------------- NIVELES ORGANIZACIONALES
                Me.ComboBox1.Enabled = True        'NIVEL 1
                Me.ComboBox2.Enabled = True        'NIVEL 2
                Me.CheckBox1.Enabled = True        'CHECK N2
                Me.ComboBox3.Enabled = True        'NIVEL 3
                Me.CheckBox2.Enabled = True        'CHECK N3
                Me.ComboBox4.Enabled = True        'NIVEL 4
                Me.CheckBox3.Enabled = True        'CHECK N4
                Me.ComboBox5.Enabled = True        'NIVEL 5
                Me.CheckBox4.Enabled = True        'CHECK N5
                Me.ComboBox6.Enabled = True        'NIVEL 6
                Me.CheckBox5.Enabled = True        'CHECK N6
                Me.ComboBox7.Enabled = True        'NIVEL 7
                Me.CheckBox6.Enabled = True        'CHECK N7
                Me.ComboBox8.Enabled = True        'NIVEL 8
                Me.CheckBox7.Enabled = True        'CHECK N8
                Me.Button1.Enabled = True          'BOTON NIVELES ORGANIZACIONALES
                '--------------- TIPO DE ESTRUCTURAS
                Me.RadioButton1.Enabled = True     'ESTRUCTURA MIXTA
                Me.RadioButton2.Enabled = True     'ESTRUCTURA MILITAR
                Me.RadioButton3.Enabled = True     'ESTRUCTURA PAME
                Me.RadioButton4.Enabled = True     'JEFE
                Me.Button2.Enabled = True          'BOTON TIPO DE ESTRUCTURAS
                '--------------- GRADOS, CARGOS, OTROS
                Me.ComboBox9.Enabled = True        'GRADO REAL
                Me.Button3.Enabled = True          'BOTON GR
                Me.ComboBox10.Enabled = True       'GRADO PLANTILLA
                Me.Button4.Enabled = True          'BOTON GP
                Me.ComboBox11.Enabled = True       'SITUACION DEL CARGO
                Me.Button5.Enabled = True          'BOTON SITUACION CARGO
                Me.ComboBox12.Enabled = True       'CARGO
                Me.Button6.Enabled = True          'BOTON CARGO
                Me.ComboBox13.Enabled = True       'ESTABLECIMIENTO
                Me.Button7.Enabled = True          'BOTON ESTABLECIMIENTO
                '--------------- FECHAS
                'Me.RadioButton5.Enabled = True    'FECHA DE INGRESO PAME
                'Me.RadioButton5.Checked = True
                'Me.RadioButton6.Enabled = True    'FECHA DE INGRESO EN
                'Me.RadioButton7.Enabled = True    'FECHA DE NACIMIENTO
                'Me.RadioButton7.Checked = True
                'Me.RadioButton8.Enabled = True    'FECHA DE SITUACION
                'Me.RadioButton9.Enabled = True    'FECHA DE MARCA
                'Me.RadioButton9.Checked = True
                'Me.RadioButton10.Enabled = True   'FECHA DE PROCESO DEL ASPIRANTE
                'Me.RadioButton11.Enabled = True   'FECHA DE APLICACION DE MOVIMIENTOS
                'Me.RadioButton15.Enabled = True   'FECHA DE EVENTUALIDADES
                'Me.RadioButton16.Enabled = True   'FECHA DE SUBSIDIOS
                'Me.RadioButton16.Checked = True
                'Me.DateTimePicker1.Enabled = True 'DESDE
                'Me.DateTimePicker2.Enabled = True 'HASTA
                'Me.Button11.Enabled = True        'BOTON FECHAS
                '--------------- OTROS FILTROS
                'Me.ComboBox19.Enabled = True      'FUNCIONALIDAD DE EMPLEADO
                'Me.Button14.Enabled = True        'BOTON FUNCIONALIDAD DE EMPLEADO
                'Me.ComboBox20.Enabled = True      'CAUSA REAL DE EVENTUALIDAD
                'Me.Button15.Enabled = True        'BOTON CAUSA REAL DE EVENTUALIDAD
                'Me.ComboBox21.Enabled = True      'TIPO DE EVENTUALIDAD
                'Me.Button16.Enabled = True        'BOTON TIPO DE EVENTUALIDAD
                'Me.ComboBox22.Enabled = True      'ESTUDIOS
                'Me.Button17.Enabled = True        'BOTON ESTUDIOS
                'Me.ComboBox23.Enabled = True      'NIVEL DE COMPETENCIA
                'Me.Button18.Enabled = True        'BOTON NIVEL DE COMPETENCIA
                'Me.ComboBox24.Enabled = True      'RELOJ DE MARCACION
                'Me.Button19.Enabled = True        'BOTON RELOJ DE MARCACION
                'Me.ComboBox25.Enabled = True      'TIPO DE MARCA
                'Me.Button20.Enabled = True        'BOTON TIPO DE MARCA
                'Me.ComboBox26.Enabled = True      'CLASIFICACION DEL EMPLEADO
                'Me.Button21.Enabled = True        'BOTON CLASIFICACION
                'Me.TextBox1.Enabled = True        'CODIGO DEL EMPLEADO
                'Me.Button22.Enabled = True        'BOTON CODIGO EMPLEADO
                'Me.RadioButton12.Enabled = True   'EDAD
                'Me.RadioButton13.Enabled = True   'ANTIGUEDAD PAME
                'Me.RadioButton14.Enabled = True   'ANTIGUEDAD EN
                'Me.TextBox2.Enabled = True        'ANTIGUEDADES A
                'Me.TextBox3.Enabled = True        'ANTIGUEDADES B
                'Me.Button23.Enabled = True        'BOTON ANTIGUEDADES
                'Me.TextBox4.Enabled = True        'CONTRATOS
                'Me.Button24.Enabled = True        'BOTON CONTRATOS
                '--------------- CONDECORACIONES
                'Me.ComboBox27.Enabled = True      'CLASIFICACION
                'Me.Button25.Enabled = True        'BOTON CLASIFICACION
                'Me.ComboBox28.Enabled = True      'TIPO
                'Me.Button26.Enabled = True        'BOTON TIPO
                '--------------- SITUACIONES
                Me.ComboBox29.Enabled = True      'TIPO SITUACION
                Me.ComboBox29.Text = "AUSENTE"
                Me.Button27.Enabled = True        'BOTON TIPO SITUACION
                Me.ComboBox30.Enabled = True      'SITUACION
                Me.ComboBox30.Text = "VACACIONES"
                Me.Button28.Enabled = True        'BOTON SITUACION
                Me.DateTimePicker3.Enabled = True  'DESDE
                Me.DateTimePicker4.Enabled = True  'HASTA
                Me.Button29.Enabled = True         'BOTON FECHAS
                ULTIMA_LISTA_QUE_HIZO_CLIC = Mid(CODIGO_REPORTE, 1, 1)
                GoTo SALIR
            End If
SALIR:
        End If
    End Sub
    Private Sub ListBox8_Click(sender As Object, e As EventArgs) Handles ListBox8.Click
        If Me.ListBox8.Text <> "" Then
            Call DESACTIVAR_TODO()
            CODIGO_REPORTE = Mid(Me.ListBox8.Text, 1, 5)
            If Mid(CODIGO_REPORTE, 1, 1) <> ULTIMA_LISTA_QUE_HIZO_CLIC Then
                Call LIMPIAR_LISTAS_DE_SELECCION()
                Me.Lista_Principal.Items.Clear()
            End If
            NOMBRE_REPORTE = Me.ListBox8.Text
            Label2.Text = NOMBRE_REPORTE
            Label2.Text = "Higiene y Seguridad\" & NOMBRE_REPORTE

            ULTIMA_LISTA_QUE_HIZO_CLIC = Mid(CODIGO_REPORTE, 1, 1)
        End If
    End Sub
    'NIVELES ORGANIZACIONALES
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.ComboBox1.SelectedValue <> 0 Then         'N1
            Me.Lista_Principal.Items.Add("Nivel 1: " & Me.ComboBox1.Text)
            If Mid(CODIGO_REPORTE, 1, 1) = "M" Then
                If Me.Lista101.Items.Count <> 0 Then
                    Me.Lista101.Items.Add(" OR ")
                End If
                Me.Lista101.Items.Add("({HISTORICO_DE_MOVIMIENTO.D_N1} = " & Me.ComboBox1.Text & ")")
                GoTo SEGUNDO_NIVEL
            End If
            If CODIGO_REPORTE = "E0005" Then
                If Me.Lista101.Items.Count <> 0 Then
                    Me.Lista101.Items.Add(" OR ")
                End If
                Me.Lista101.Items.Add("({VISTA_CONSULTA_MAESTRO_DE_CURVA_Y_TALLA.ID_N1} = " & Me.ComboBox1.SelectedValue & ")")
                GoTo SEGUNDO_NIVEL
            End If
            If CODIGO_REPORTE = "S0004" Then
                If Me.Lista101.Items.Count <> 0 Then
                    Me.Lista101.Items.Add(" OR ")
                End If
                Me.Lista101.Items.Add("({VISTA_MAESTRO_DE_CARGOS_SIT.ID_N1} = " & Me.ComboBox1.SelectedValue & ")")
                GoTo SEGUNDO_NIVEL
            End If
            If Me.Lista101.Items.Count <> 0 Then
                Me.Lista101.Items.Add(" OR ")
            End If
            Me.Lista101.Items.Add("({MAESTRO_DE_CARGOS.ID_N1} = " & Me.ComboBox1.SelectedValue & ")")
        Else
            MsgBox("La Seleccion para el Nivel 1 no es válida", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
SEGUNDO_NIVEL:
        If Me.CheckBox1.Checked = True Then             'N2
            If Me.ComboBox2.SelectedValue <> 0 Then
                Me.Lista_Principal.Items.Add("Nivel 2: " & Me.ComboBox2.Text)
                If Mid(CODIGO_REPORTE, 1, 1) = "M" Then
                    If Me.Lista102.Items.Count <> 0 Then
                        Me.Lista102.Items.Add(" OR ")
                    End If
                    Me.Lista102.Items.Add("({HISTORICO_DE_MOVIMIENTO.D_N2} = " & Me.ComboBox2.Text & ")")
                    GoTo TERCER_NIVEL
                End If
                If CODIGO_REPORTE = "E0005" Then
                    If Me.Lista102.Items.Count <> 0 Then
                        Me.Lista102.Items.Add(" OR ")
                    End If
                    Me.Lista102.Items.Add("({VISTA_CONSULTA_MAESTRO_DE_CURVA_Y_TALLA.ID_N2} = " & Me.ComboBox2.SelectedValue & ")")
                    GoTo TERCER_NIVEL
                End If
                If CODIGO_REPORTE = "S0004" Then
                    If Me.Lista102.Items.Count <> 0 Then
                        Me.Lista102.Items.Add(" OR ")
                    End If
                    Me.Lista102.Items.Add("({VISTA_MAESTRO_DE_CARGOS_SIT.ID_N2} = " & Me.ComboBox2.SelectedValue & ")")
                    GoTo TERCER_NIVEL
                End If
                If Me.Lista102.Items.Count <> 0 Then
                    Me.Lista102.Items.Add(" OR ")
                End If
                Me.Lista102.Items.Add("({MAESTRO_DE_CARGOS.ID_N2} = " & Me.ComboBox2.SelectedValue & ")")
            Else
                MsgBox("La Seleccion para el Nivel 2 no es válida", vbInformation, "Mensaje del Sistema")
                Me.ComboBox2.Focus()
                Exit Sub
            End If
        End If
TERCER_NIVEL:
        If Me.CheckBox2.Checked = True Then             'N3
            If Me.ComboBox3.SelectedValue <> 0 Then
                Me.Lista_Principal.Items.Add("Nivel 3: " & Me.ComboBox3.Text)
                If Mid(CODIGO_REPORTE, 1, 1) = "M" Then
                    If Me.Lista103.Items.Count <> 0 Then
                        Me.Lista103.Items.Add(" OR ")
                    End If
                    Me.Lista103.Items.Add("({HISTORICO_DE_MOVIMIENTO.D_N3} = " & Me.ComboBox3.Text & ")")
                    GoTo CUARTO_NIVEL
                End If
                If CODIGO_REPORTE = "E0005" Then
                    If Me.Lista103.Items.Count <> 0 Then
                        Me.Lista103.Items.Add(" OR ")
                    End If
                    Me.Lista103.Items.Add("({VISTA_CONSULTA_MAESTRO_DE_CURVA_Y_TALLA.ID_N3} = " & Me.ComboBox3.SelectedValue & ")")
                    GoTo CUARTO_NIVEL
                End If
                If CODIGO_REPORTE = "S0004" Then
                    If Me.Lista103.Items.Count <> 0 Then
                        Me.Lista103.Items.Add(" OR ")
                    End If
                    Me.Lista103.Items.Add("({VISTA_MAESTRO_DE_CARGOS_SIT.ID_N3} = " & Me.ComboBox3.SelectedValue & ")")
                    GoTo CUARTO_NIVEL
                End If
                If Me.Lista103.Items.Count <> 0 Then
                    Me.Lista103.Items.Add(" OR ")
                End If
                Me.Lista103.Items.Add("({MAESTRO_DE_CARGOS.ID_N3} = " & Me.ComboBox3.SelectedValue & ")")
            Else
                MsgBox("La Seleccion para el Nivel 3 no es válida", vbInformation, "Mensaje del Sistema")
                Me.ComboBox3.Focus()
                Exit Sub
            End If
        End If
CUARTO_NIVEL:
        If Me.CheckBox3.Checked = True Then             'N4
            If Me.ComboBox4.SelectedValue <> 0 Then
                Me.Lista_Principal.Items.Add("Nivel 4: " & Me.ComboBox4.Text)
                If Mid(CODIGO_REPORTE, 1, 1) = "M" Then
                    If Me.Lista104.Items.Count <> 0 Then
                        Me.Lista104.Items.Add(" OR ")
                    End If
                    Me.Lista104.Items.Add("({HISTORICO_DE_MOVIMIENTO.D_N4} = " & Me.ComboBox4.Text & ")")
                    GoTo QUINTO_NIVEL
                End If
                If CODIGO_REPORTE = "E0005" Then
                    If Me.Lista104.Items.Count <> 0 Then
                        Me.Lista104.Items.Add(" OR ")
                    End If
                    Me.Lista104.Items.Add("({VISTA_CONSULTA_MAESTRO_DE_CURVA_Y_TALLA.ID_N4} = " & Me.ComboBox4.SelectedValue & ")")
                    GoTo QUINTO_NIVEL
                End If
                If CODIGO_REPORTE = "S0004" Then
                    If Me.Lista104.Items.Count <> 0 Then
                        Me.Lista104.Items.Add(" OR ")
                    End If
                    Me.Lista104.Items.Add("({VISTA_MAESTRO_DE_CARGOS_SIT.ID_N4} = " & Me.ComboBox4.SelectedValue & ")")
                    GoTo QUINTO_NIVEL
                End If
                '''''''''''''''''''''''''''''''
                If Me.Lista104.Items.Count <> 0 Then
                    Me.Lista104.Items.Add(" OR ")
                End If
                Me.Lista104.Items.Add("({MAESTRO_DE_CARGOS.ID_N4} = " & Me.ComboBox4.SelectedValue & ")")
            Else
                MsgBox("La Seleccion para el Nivel 4 no es válida", vbInformation, "Mensaje del Sistema")
                Me.ComboBox4.Focus()
                Exit Sub
            End If
        End If
QUINTO_NIVEL:
        If Me.CheckBox4.Checked = True Then             'N5
            If Me.ComboBox5.SelectedValue <> 0 Then
                Me.Lista_Principal.Items.Add("Nivel 5: " & Me.ComboBox5.Text)
                If Mid(CODIGO_REPORTE, 1, 1) = "M" Then
                    If Me.Lista105.Items.Count <> 0 Then
                        Me.Lista105.Items.Add(" OR ")
                    End If
                    Me.Lista105.Items.Add("({HISTORICO_DE_MOVIMIENTO.D_N5} = " & Me.ComboBox5.Text & ")")
                    GoTo SEXTO_NIVEL
                End If
                If CODIGO_REPORTE = "E0005" Then
                    If Me.Lista105.Items.Count <> 0 Then
                        Me.Lista105.Items.Add(" OR ")
                    End If
                    Me.Lista105.Items.Add("({VISTA_CONSULTA_MAESTRO_DE_CURVA_Y_TALLA.ID_N5} = " & Me.ComboBox5.SelectedValue & ")")
                    GoTo SEXTO_NIVEL
                End If
                If CODIGO_REPORTE = "S0004" Then
                    If Me.Lista105.Items.Count <> 0 Then
                        Me.Lista105.Items.Add(" OR ")
                    End If
                    Me.Lista105.Items.Add("({VISTA_MAESTRO_DE_CARGOS_SIT.ID_N5} = " & Me.ComboBox5.SelectedValue & ")")
                    GoTo SEXTO_NIVEL
                End If
                '''''''''''''''''''''''''''''''
                If Me.Lista105.Items.Count <> 0 Then
                    Me.Lista105.Items.Add(" OR ")
                End If
                Me.Lista105.Items.Add("({MAESTRO_DE_CARGOS.ID_N5} = " & Me.ComboBox5.SelectedValue & ")")
            Else
                MsgBox("La Seleccion para el Nivel 5 no es válida", vbInformation, "Mensaje del Sistema")
                Me.ComboBox5.Focus()
                Exit Sub
            End If
        End If
SEXTO_NIVEL:
        If Me.CheckBox5.Checked = True Then             'N6
            If Me.ComboBox6.SelectedValue <> 0 Then
                Me.Lista_Principal.Items.Add("Nivel 6: " & Me.ComboBox6.Text)
                If Mid(CODIGO_REPORTE, 1, 1) = "M" Then
                    If Me.Lista106.Items.Count <> 0 Then
                        Me.Lista106.Items.Add(" OR ")
                    End If
                    Me.Lista106.Items.Add("({HISTORICO_DE_MOVIMIENTO.D_N6} = " & Me.ComboBox6.Text & ")")
                    GoTo SEPTIMO_NIVEL
                End If
                If CODIGO_REPORTE = "E0005" Then
                    If Me.Lista106.Items.Count <> 0 Then
                        Me.Lista106.Items.Add(" OR ")
                    End If
                    Me.Lista106.Items.Add("({VISTA_CONSULTA_MAESTRO_DE_CURVA_Y_TALLA.ID_N6} = " & Me.ComboBox6.SelectedValue & ")")
                    GoTo SEPTIMO_NIVEL
                End If
                If CODIGO_REPORTE = "S0004" Then
                    If Me.Lista106.Items.Count <> 0 Then
                        Me.Lista106.Items.Add(" OR ")
                    End If
                    Me.Lista106.Items.Add("({VISTA_MAESTRO_DE_CARGOS_SIT.ID_N6} = " & Me.ComboBox6.SelectedValue & ")")
                    GoTo SEPTIMO_NIVEL
                End If
                '''''''''''''''''''''''''''''''
                If Me.Lista106.Items.Count <> 0 Then
                    Me.Lista106.Items.Add(" OR ")
                End If
                Me.Lista106.Items.Add("({MAESTRO_DE_CARGOS.ID_N6} = " & Me.ComboBox6.SelectedValue & ")")
            Else
                MsgBox("La Seleccion para el Nivel 6 no es válida", vbInformation, "Mensaje del Sistema")
                Me.ComboBox6.Focus()
                Exit Sub
            End If
        End If
SEPTIMO_NIVEL:
        If Me.CheckBox6.Checked = True Then             'N7
            If Me.ComboBox7.SelectedValue <> 0 Then
                Me.Lista_Principal.Items.Add("Nivel 7: " & Me.ComboBox7.Text)
                If Mid(CODIGO_REPORTE, 1, 1) = "M" Then
                    If Me.Lista107.Items.Count <> 0 Then
                        Me.Lista107.Items.Add(" OR ")
                    End If
                    Me.Lista107.Items.Add("({HISTORICO_DE_MOVIMIENTO.D_N7} = " & Me.ComboBox7.Text & ")")
                    GoTo OCTAVO_NIVEL
                End If
                If CODIGO_REPORTE = "E0005" Then
                    If Me.Lista107.Items.Count <> 0 Then
                        Me.Lista107.Items.Add(" OR ")
                    End If
                    Me.Lista107.Items.Add("({VISTA_CONSULTA_MAESTRO_DE_CURVA_Y_TALLA.ID_N7} = " & Me.ComboBox7.SelectedValue & ")")
                    GoTo OCTAVO_NIVEL
                End If
                If CODIGO_REPORTE = "S0004" Then
                    If Me.Lista107.Items.Count <> 0 Then
                        Me.Lista107.Items.Add(" OR ")
                    End If
                    Me.Lista107.Items.Add("({VISTA_MAESTRO_DE_CARGOS_SIT.ID_N7} = " & Me.ComboBox7.SelectedValue & ")")
                    GoTo OCTAVO_NIVEL
                End If
                '''''''''''''''''''''''''''''''
                If Me.Lista107.Items.Count <> 0 Then
                    Me.Lista107.Items.Add(" OR ")
                End If
                Me.Lista107.Items.Add("({MAESTRO_DE_CARGOS.ID_N7} = " & Me.ComboBox7.SelectedValue & ")")
            Else
                MsgBox("La Seleccion para el Nivel 7 no es válida", vbInformation, "Mensaje del Sistema")
                Me.ComboBox7.Focus()
                Exit Sub
            End If
        End If
OCTAVO_NIVEL:
        If Me.CheckBox7.Checked = True Then             'N8
            If Me.ComboBox8.SelectedValue <> 0 Then
                Me.Lista_Principal.Items.Add("Nivel 8: " & Me.ComboBox8.Text)
                If Mid(CODIGO_REPORTE, 1, 1) = "M" Then
                    If Me.Lista108.Items.Count <> 0 Then
                        Me.Lista108.Items.Add(" OR ")
                    End If
                    Me.Lista108.Items.Add("({HISTORICO_DE_MOVIMIENTO.D_N8} = " & Me.ComboBox8.Text & ")")
                    GoTo SALIR
                End If
                If CODIGO_REPORTE = "E0005" Then
                    If Me.Lista108.Items.Count <> 0 Then
                        Me.Lista108.Items.Add(" OR ")
                    End If
                    Me.Lista108.Items.Add("({VISTA_CONSULTA_MAESTRO_DE_CURVA_Y_TALLA.ID_N8} = " & Me.ComboBox8.SelectedValue & ")")
                    GoTo SALIR
                End If
                If CODIGO_REPORTE = "S0004" Then
                    If Me.Lista108.Items.Count <> 0 Then
                        Me.Lista108.Items.Add(" OR ")
                    End If
                    Me.Lista108.Items.Add("({VISTA_MAESTRO_DE_CARGOS_SIT.ID_N8} = " & Me.ComboBox8.SelectedValue & ")")
                    GoTo SALIR
                End If
                '''''''''''''''''''''''''''''''
                If Me.Lista108.Items.Count <> 0 Then
                    Me.Lista108.Items.Add(" OR ")
                End If
                Me.Lista108.Items.Add("({MAESTRO_DE_CARGOS.ID_N8} = " & Me.ComboBox8.SelectedValue & ")")
            Else
                MsgBox("La Seleccion para el Nivel 8 no es válida", vbInformation, "Mensaje del Sistema")
                Me.ComboBox8.Focus()
                Exit Sub
            End If
        End If
SALIR:
    End Sub
    'TIPO DE ESTRUCTURA
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Me.RadioButton1.Checked = False And Me.RadioButton2.Checked = False And Me.RadioButton3.Checked = False And Me.RadioButton4.Checked = False Then
            MsgBox("La Seleccion para el Tipo de Estructura no es válida", vbInformation, "Mensaje del Sistema")
            Me.RadioButton3.Focus()
            Exit Sub
        End If
        If Me.RadioButton1.Checked = True Then         'MIXTA
            Me.Lista_Principal.Items.Add("Tipo de Estructura: MIXTA")
            If Mid(CODIGO_REPORTE, 1, 1) = "M" Then
                If Me.Lista109.Items.Count <> 0 Then
                    Me.Lista109.Items.Add(" OR ")
                End If
                Me.Lista109.Items.Add("({HISTORICO_DE_MOVIMIENTO.MIXTA} = TRUE)")
                GoTo SALIR
            End If
            If CODIGO_REPORTE = "S0004" Then
                If Me.Lista109.Items.Count <> 0 Then
                    Me.Lista109.Items.Add(" OR ")
                End If
                Me.Lista109.Items.Add("({VISTA_MAESTRO_DE_CARGOS_SIT.MIXTA} = TRUE)")
                GoTo SALIR
            End If
            '''''''''''''''''''''''''''''''
            If Me.Lista109.Items.Count <> 0 Then
                Me.Lista109.Items.Add(" OR ")
            End If
            Me.Lista109.Items.Add("({MAESTRO_DE_CARGOS.MIXTA} = TRUE)")
        End If

        If Me.RadioButton2.Checked = True Then         'MILITAR
            Me.Lista_Principal.Items.Add("Tipo de Estructura: MILITAR")
            If Mid(CODIGO_REPORTE, 1, 1) = "M" Then
                If Me.Lista109.Items.Count <> 0 Then
                    Me.Lista109.Items.Add(" OR ")
                End If
                Me.Lista109.Items.Add("({HISTORICO_DE_MOVIMIENTO.MILITAR} = TRUE)")
                GoTo SALIR
            End If
            If CODIGO_REPORTE = "S0004" Then
                If Me.Lista109.Items.Count <> 0 Then
                    Me.Lista109.Items.Add(" OR ")
                End If
                Me.Lista109.Items.Add("({VISTA_MAESTRO_DE_CARGOS_SIT.MILITAR} = TRUE)")
                GoTo SALIR
            End If
            '''''''''''''''''''''''''''''''
            If Me.Lista109.Items.Count <> 0 Then
                Me.Lista109.Items.Add(" OR ")
            End If
            Me.Lista109.Items.Add("({MAESTRO_DE_CARGOS.MILITAR} = TRUE)")
        End If

        If Me.RadioButton3.Checked = True Then         'PAME
            Me.Lista_Principal.Items.Add("Tipo de Estructura: PAME")
            If Mid(CODIGO_REPORTE, 1, 1) = "M" Then
                If Me.Lista109.Items.Count <> 0 Then
                    Me.Lista109.Items.Add(" OR ")
                End If
                Me.Lista109.Items.Add("({HISTORICO_DE_MOVIMIENTO.PAME} = TRUE)")
                GoTo SALIR
            End If
            If CODIGO_REPORTE = "S0004" Then
                If Me.Lista109.Items.Count <> 0 Then
                    Me.Lista109.Items.Add(" OR ")
                End If
                Me.Lista109.Items.Add("({VISTA_MAESTRO_DE_CARGOS_SIT.PAME} = TRUE)")
                GoTo SALIR
            End If
            '''''''''''''''''''''''''''''''
            If Me.Lista109.Items.Count <> 0 Then
                Me.Lista109.Items.Add(" OR ")
            End If
            Me.Lista109.Items.Add("({MAESTRO_DE_CARGOS.PAME} = TRUE)")
        End If

        If Me.RadioButton4.Checked = True Then         'JEFE
            Me.Lista_Principal.Items.Add("Tipo de Estructura: JEFE")
            If Mid(CODIGO_REPORTE, 1, 1) = "M" Then
                If Me.Lista109.Items.Count <> 0 Then
                    Me.Lista109.Items.Add(" OR ")
                End If
                Me.Lista109.Items.Add("({HISTORICO_DE_MOVIMIENTO.JEFE} = TRUE)")
                GoTo SALIR
            End If
            If CODIGO_REPORTE = "S0004" Then
                If Me.Lista109.Items.Count <> 0 Then
                    Me.Lista109.Items.Add(" OR ")
                End If
                Me.Lista109.Items.Add("({VISTA_MAESTRO_DE_CARGOS_SIT.JEFE} = TRUE)")
                GoTo SALIR
            End If
            '''''''''''''''''''''''''''''''
            If Me.Lista109.Items.Count <> 0 Then
                Me.Lista109.Items.Add(" OR ")
            End If
            Me.Lista109.Items.Add("({MAESTRO_DE_CARGOS.JEFE} = TRUE)")
        End If
SALIR:
    End Sub
    'GRADO REAL
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Me.ComboBox9.SelectedValue <> 0 Then
            Me.Lista_Principal.Items.Add("Grado Real: " & Me.ComboBox9.Text)
            If Mid(CODIGO_REPORTE, 1, 1) = "M" Then
                If Me.Lista110.Items.Count <> 0 Then
                    Me.Lista110.Items.Add(" OR ")
                End If
                Me.Lista110.Items.Add("({HISTORICO_DE_MOVIMIENTO.D_GR} = '" & Me.ComboBox9.Text & "')")
                GoTo SALIR
            End If
            If CODIGO_REPORTE = "E0005" Then
                If Me.Lista110.Items.Count <> 0 Then
                    Me.Lista110.Items.Add(" OR ")
                End If
                Me.Lista110.Items.Add("({VISTA_CONSULTA_MAESTRO_DE_CURVA_Y_TALLA.ID_GR} = " & Me.ComboBox9.SelectedValue & ")")
                GoTo SALIR
            End If
            If CODIGO_REPORTE = "S0004" Then
                If Me.Lista110.Items.Count <> 0 Then
                    Me.Lista110.Items.Add(" OR ")
                End If
                Me.Lista110.Items.Add("({VISTA_MAESTRO_DE_CARGOS_SIT.ID_GR} = " & Me.ComboBox9.SelectedValue & ")")
                GoTo SALIR
            End If
            '''''''''''''''''''''''''''''''
            If Me.Lista110.Items.Count <> 0 Then
                Me.Lista110.Items.Add(" OR ")
            End If
            Me.Lista110.Items.Add("({MAESTRO_DE_CARGOS.ID_GR} = " & Me.ComboBox9.SelectedValue & ")")
        Else
            MsgBox("La Seleccion para el Grado Real no es válida", vbInformation, "Mensaje del Sistema")
            Me.ComboBox9.Focus()
            Exit Sub
        End If
SALIR:
    End Sub
    'GRADO PLANTILLA
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If Me.ComboBox10.SelectedValue <> 0 Then
            Me.Lista_Principal.Items.Add("Grado Plantilla: " & Me.ComboBox10.Text)
            If Mid(CODIGO_REPORTE, 1, 1) = "M" Then
                If Me.Lista111.Items.Count <> 0 Then
                    Me.Lista111.Items.Add(" OR ")
                End If
                Me.Lista111.Items.Add("({HISTORICO_DE_MOVIMIENTO.D_GP} = '" & Me.ComboBox10.Text & "')")
                GoTo SALIR
            End If
            If CODIGO_REPORTE = "E0005" Then
                If Me.Lista111.Items.Count <> 0 Then
                    Me.Lista111.Items.Add(" OR ")
                End If
                Me.Lista111.Items.Add("({VISTA_CONSULTA_MAESTRO_DE_CURVA_Y_TALLA.ID_GP} = " & Me.ComboBox10.SelectedValue & ")")
                GoTo SALIR
            End If
            If CODIGO_REPORTE = "S0004" Then
                If Me.Lista111.Items.Count <> 0 Then
                    Me.Lista111.Items.Add(" OR ")
                End If
                Me.Lista111.Items.Add("({VISTA_MAESTRO_DE_CARGOS_SIT.ID_GP} = " & Me.ComboBox10.SelectedValue & ")")
                GoTo SALIR
            End If
            If Me.Lista111.Items.Count <> 0 Then
                Me.Lista111.Items.Add(" OR ")
            End If
            Me.Lista111.Items.Add("({MAESTRO_DE_CARGOS.ID_GP} = " & Me.ComboBox10.SelectedValue & ")")
        Else
            MsgBox("La Seleccion para el Grado Plantilla no es válida", vbInformation, "Mensaje del Sistema")
            Me.ComboBox10.Focus()
            Exit Sub
        End If
SALIR:
    End Sub
    'SITUACION DEL CARGO
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If Me.ComboBox11.SelectedValue <> 0 Then
            Me.Lista_Principal.Items.Add("Situacion: " & Me.ComboBox11.Text)
            If CODIGO_REPORTE = "S0004" Then
                If Me.Lista112.Items.Count <> 0 Then
                    Me.Lista112.Items.Add(" OR ")
                End If
                Me.Lista112.Items.Add("({VISTA_MAESTRO_DE_CARGOS_SIT.ID_SITUACION} = " & Me.ComboBox11.SelectedValue & ")")
                GoTo SALIR
            End If
            '''''''''''''''''''''''''''''''
            If Me.Lista112.Items.Count <> 0 Then
                Me.Lista112.Items.Add(" OR ")
            End If
            Me.Lista112.Items.Add("({MAESTRO_DE_CARGOS.ID_SITUACION} = " & Me.ComboBox11.SelectedValue & ")")
        Else
            MsgBox("La Seleccion para el Tipo de Situacion no es válida", vbInformation, "Mensaje del Sistema")
            Me.ComboBox11.Focus()
            Exit Sub
        End If
SALIR:
    End Sub
    'CARGO
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If Me.ComboBox12.SelectedValue <> 0 Then
            Me.Lista_Principal.Items.Add("Cargo: " & Me.ComboBox12.Text)
            If Mid(CODIGO_REPORTE, 1, 1) = "M" Then
                If Me.Lista113.Items.Count <> 0 Then
                    Me.Lista113.Items.Add(" OR ")
                End If
                Me.Lista113.Items.Add("({HISTORICO_DE_MOVIMIENTO.D_CARGO_ES} = '" & Me.ComboBox12.Text & "')")
                GoTo SALIR
            End If
            If CODIGO_REPORTE = "E0005" Then
                If Me.Lista113.Items.Count <> 0 Then
                    Me.Lista113.Items.Add(" OR ")
                End If
                Me.Lista113.Items.Add("({VISTA_CONSULTA_MAESTRO_DE_CURVA_Y_TALLA.ID_CARGO_ES} = " & Me.ComboBox12.SelectedValue & ")")
                GoTo SALIR
            End If
            If CODIGO_REPORTE = "S0004" Then
                If Me.Lista113.Items.Count <> 0 Then
                    Me.Lista113.Items.Add(" OR ")
                End If
                Me.Lista113.Items.Add("({VISTA_MAESTRO_DE_CARGOS_SIT.ID_CARGO_ES} = " & Me.ComboBox12.SelectedValue & ")")
                GoTo SALIR
            End If
            If Me.Lista113.Items.Count <> 0 Then
                Me.Lista113.Items.Add(" OR ")
            End If
            Me.Lista113.Items.Add("({MAESTRO_DE_CARGOS.ID_CARGO_ES} = " & Me.ComboBox12.SelectedValue & ")")
        Else
            MsgBox("La Seleccion para el Cargo no es válida", vbInformation, "Mensaje del Sistema")
            Me.ComboBox12.Focus()
            Exit Sub
        End If
SALIR:
    End Sub
    'ESTABLECIMIENTO
    Private Sub Button7_Click_1(sender As Object, e As EventArgs) Handles Button7.Click
        If Me.ComboBox13.SelectedValue <> 0 Then
            Me.Lista_Principal.Items.Add("Establecimiento: " & Me.ComboBox13.Text)
            If CODIGO_REPORTE = "E0005" Then
                If Me.Lista114.Items.Count <> 0 Then
                    Me.Lista114.Items.Add(" OR ")
                End If
                Me.Lista114.Items.Add("({VISTA_CONSULTA_MAESTRO_DE_CURVA_Y_TALLA.ID_ESTABLECIMIENTO} = " & Me.ComboBox13.SelectedValue & ")")
                GoTo SALIR
            End If
            If CODIGO_REPORTE = "S0004" Then
                If Me.Lista114.Items.Count <> 0 Then
                    Me.Lista114.Items.Add(" OR ")
                End If
                Me.Lista114.Items.Add("({VISTA_MAESTRO_DE_CARGOS_SIT.ID_ESTABLECIMIENTO} = " & Me.ComboBox13.SelectedValue & ")")
                GoTo SALIR
            End If
            If Me.Lista114.Items.Count <> 0 Then
                Me.Lista114.Items.Add(" OR ")
            End If
            Me.Lista114.Items.Add("({MAESTRO_DE_CARGOS.ID_ESTABLECIMIENTO} = " & Me.ComboBox13.SelectedValue & ")")
        Else
            MsgBox("La Seleccion para el Establecimiento no es válida", vbInformation, "Mensaje del Sistema")
            Me.ComboBox13.Focus()
            Exit Sub
        End If
SALIR:
    End Sub
    'TIPO DE MOVIMIENTO
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If Me.ComboBox14.SelectedValue <> 0 Then
            Me.Lista_Principal.Items.Add("Tipo de Movimiento: " & Me.ComboBox14.Text)
            If Me.Lista115.Items.Count = 0 Then
                Me.Lista115.Items.Add("({HISTORICO_DE_MOVIMIENTO.ID_T_M_EXP} = " & Me.ComboBox14.SelectedValue & ")")
            Else
                Me.Lista115.Items.Add(" OR ")
                Me.Lista115.Items.Add("({HISTORICO_DE_MOVIMIENTO.ID_T_M_EXP} = " & Me.ComboBox14.SelectedValue & ")")
            End If
        Else
            MsgBox("La Seleccion para el Tipo de Movimiento no es válida", vbInformation, "Mensaje del Sistema")
            Me.ComboBox14.Focus()
            Exit Sub
        End If
    End Sub
    'CAUSA REAL 
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If Me.ComboBox15.SelectedValue <> 0 Then
            Me.Lista_Principal.Items.Add("Causa Real del Movimiento: " & Me.ComboBox15.Text)
            If Me.Lista116.Items.Count = 0 Then
                Me.Lista116.Items.Add("({HISTORICO_DE_MOVIMIENTO.ID_ST_M_R_EXP} = " & Me.ComboBox15.SelectedValue & ")")
            Else
                Me.Lista116.Items.Add(" OR ")
                Me.Lista116.Items.Add("({HISTORICO_DE_MOVIMIENTO.ID_ST_M_R_EXP} = " & Me.ComboBox15.SelectedValue & ")")
            End If
        Else
            MsgBox("La Seleccion para la Causa Real del Movimiento no es válida", vbInformation, "Mensaje del Sistema")
            Me.ComboBox15.Focus()
            Exit Sub
        End If
    End Sub
    'CAUSA LEGAL
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        If Me.ComboBox16.SelectedValue <> 0 Then
            Me.Lista_Principal.Items.Add("Causa Legal del Movimiento: " & Me.ComboBox16.Text)
            If Me.Lista117.Items.Count = 0 Then
                Me.Lista117.Items.Add("({HISTORICO_DE_MOVIMIENTO.ID_ST_M_L_EXP} = " & Me.ComboBox16.SelectedValue & ")")
            Else
                Me.Lista117.Items.Add(" OR ")
                Me.Lista117.Items.Add("({HISTORICO_DE_MOVIMIENTO.ID_ST_M_L_EXP} = " & Me.ComboBox16.SelectedValue & ")")
            End If
        Else
            MsgBox("La Seleccion para la Causa Legal del Movimiento no es válida", vbInformation, "Mensaje del Sistema")
            Me.ComboBox16.Focus()
            Exit Sub
        End If
    End Sub
    'FECHAS
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        If Me.DateTimePicker1.Value > Me.DateTimePicker2.Value Then
            MsgBox("La Fecha Inicial (DESDE) no puede ser mayo a la Fecha Final (HASTA)", vbInformation, "Mensaje del Sistema")
            Me.DateTimePicker1.Focus()
            Exit Sub
        End If
        If Me.RadioButton5.Checked = True Then         'FECHA INGRESO PAME
            If Mid(Me.DateTimePicker1.Value, 1, 10) = Mid(Me.DateTimePicker2.Value, 1, 10) Then
                Me.Lista_Principal.Items.Add("Fecha Ingreso PAME: " & Mid(Me.DateTimePicker1.Value, 1, 10))
            Else
                Me.Lista_Principal.Items.Add("Fecha Ingreso PAME: Del " & Mid(Me.DateTimePicker1.Value, 1, 10) & " al " & Mid(Me.DateTimePicker2.Value, 1, 10))
            End If
            D1 = Mid(Me.DateTimePicker1.Value, 1, 2)
            M1 = Mid(Me.DateTimePicker1.Value, 4, 2)
            A1 = Mid(Me.DateTimePicker1.Value, 7, 4)
            D2 = Mid(Me.DateTimePicker2.Value, 1, 2)
            M2 = Mid(Me.DateTimePicker2.Value, 4, 2)
            A2 = Mid(Me.DateTimePicker2.Value, 7, 4)
            If Mid(Me.DateTimePicker1.Value, 1, 10) = Mid(Me.DateTimePicker2.Value, 1, 10) Then
                If CODIGO_REPORTE = "E0002" Then
                    If Me.Lista118.Items.Count = 0 Then
                        Me.Lista118.Items.Add("{VISTA_MAESTRO_DE_PERSONAS.FEC_ING_PAME} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ")")
                    Else
                        Me.Lista118.Items.Add(" OR ")
                        Me.Lista118.Items.Add("{VISTA_MAESTRO_DE_PERSONAS.FEC_ING_PAME} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ")")
                    End If
                Else
                    If Me.Lista118.Items.Count = 0 Then
                        Me.Lista118.Items.Add("{MAESTRO_DE_PERSONAS.FEC_ING_PAME} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ")")
                    Else
                        Me.Lista118.Items.Add(" OR ")
                        Me.Lista118.Items.Add("{MAESTRO_DE_PERSONAS.FEC_ING_PAME} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ")")
                    End If
                End If
            Else
                If CODIGO_REPORTE = "E0002" Then
                    If Me.Lista118.Items.Count = 0 Then
                        Me.Lista118.Items.Add("{VISTA_MAESTRO_DE_PERSONAS.FEC_ING_PAME} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ") To CDateTime (" & A2 & ", " & M2 & ", " & D2 & ")")
                    Else
                        Me.Lista118.Items.Add(" OR ")
                        Me.Lista118.Items.Add("{VISTA_MAESTRO_DE_PERSONAS.FEC_ING_PAME} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ") To CDateTime (" & A2 & ", " & M2 & ", " & D2 & ")")
                    End If
                Else
                    If Me.Lista118.Items.Count = 0 Then
                        Me.Lista118.Items.Add("{MAESTRO_DE_PERSONAS.FEC_ING_PAME} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ") To CDateTime (" & A2 & ", " & M2 & ", " & D2 & ")")
                    Else
                        Me.Lista118.Items.Add(" OR ")
                        Me.Lista118.Items.Add("{MAESTRO_DE_PERSONAS.FEC_ING_PAME} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ") To CDateTime (" & A2 & ", " & M2 & ", " & D2 & ")")
                    End If
                End If
            End If
        End If
        If Me.RadioButton6.Checked = True Then         'FECHA INGRESO EN
            If Mid(Me.DateTimePicker1.Value, 1, 10) = Mid(Me.DateTimePicker2.Value, 1, 10) Then
                Me.Lista_Principal.Items.Add("Fecha Ingreso EN: " & Mid(Me.DateTimePicker1.Value, 1, 10))
            Else
                Me.Lista_Principal.Items.Add("Fecha Ingreso EN: Del " & Mid(Me.DateTimePicker1.Value, 1, 10) & " al " & Mid(Me.DateTimePicker2.Value, 1, 10))
            End If
            D1 = Mid(Me.DateTimePicker1.Value, 1, 2)
            M1 = Mid(Me.DateTimePicker1.Value, 4, 2)
            A1 = Mid(Me.DateTimePicker1.Value, 7, 4)
            D2 = Mid(Me.DateTimePicker2.Value, 1, 2)
            M2 = Mid(Me.DateTimePicker2.Value, 4, 2)
            A2 = Mid(Me.DateTimePicker2.Value, 7, 4)
            If Mid(Me.DateTimePicker1.Value, 1, 10) = Mid(Me.DateTimePicker2.Value, 1, 10) Then
                If CODIGO_REPORTE = "E0002" Then
                    If Me.Lista118.Items.Count = 0 Then
                        Me.Lista118.Items.Add("{VISTA_MAESTRO_DE_PERSONAS.FEC_ING_EN} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ")")
                    Else
                        Me.Lista118.Items.Add(" OR ")
                        Me.Lista118.Items.Add("{VISTA_MAESTRO_DE_PERSONAS.FEC_ING_EN} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ")")
                    End If
                Else
                    If Me.Lista118.Items.Count = 0 Then
                        Me.Lista118.Items.Add("{MAESTRO_DE_PERSONAS.FEC_ING_EN} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ")")
                    Else
                        Me.Lista118.Items.Add(" OR ")
                        Me.Lista118.Items.Add("{MAESTRO_DE_PERSONAS.FEC_ING_EN} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ")")
                    End If
                End If
            Else
                If CODIGO_REPORTE = "E0002" Then
                    If Me.Lista118.Items.Count = 0 Then
                        Me.Lista118.Items.Add("{VISTA_MAESTRO_DE_PERSONAS.FEC_ING_EN} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ") To CDateTime (" & A2 & ", " & M2 & ", " & D2 & ")")
                    Else
                        Me.Lista118.Items.Add(" OR ")
                        Me.Lista118.Items.Add("{VISTA_MAESTRO_DE_PERSONAS.FEC_ING_EN} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ") To CDateTime (" & A2 & ", " & M2 & ", " & D2 & ")")
                    End If
                Else
                    If Me.Lista118.Items.Count = 0 Then
                        Me.Lista118.Items.Add("{MAESTRO_DE_PERSONAS.FEC_ING_EN} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ") To CDateTime (" & A2 & ", " & M2 & ", " & D2 & ")")
                    Else
                        Me.Lista118.Items.Add(" OR ")
                        Me.Lista118.Items.Add("{MAESTRO_DE_PERSONAS.FEC_ING_EN} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ") To CDateTime (" & A2 & ", " & M2 & ", " & D2 & ")")
                    End If
                End If
            End If
        End If
        If Me.RadioButton7.Checked = True Then         'FECHA NACIMIENTO
            If Mid(Me.DateTimePicker1.Value, 1, 10) = Mid(Me.DateTimePicker2.Value, 1, 10) Then
                Me.Lista_Principal.Items.Add("Fecha Nacimiento: " & Mid(Me.DateTimePicker1.Value, 1, 10))
            Else
                Me.Lista_Principal.Items.Add("Fecha Nacimiento: Del " & Mid(Me.DateTimePicker1.Value, 1, 10) & " al " & Mid(Me.DateTimePicker2.Value, 1, 10))
            End If
            D1 = Mid(Me.DateTimePicker1.Value, 1, 2)
            M1 = Mid(Me.DateTimePicker1.Value, 4, 2)
            A1 = Mid(Me.DateTimePicker1.Value, 7, 4)
            D2 = Mid(Me.DateTimePicker2.Value, 1, 2)
            M2 = Mid(Me.DateTimePicker2.Value, 4, 2)
            A2 = Mid(Me.DateTimePicker2.Value, 7, 4)
            If CODIGO_REPORTE = "E0002" Then
                If Me.Lista118.Items.Count <> 0 Then
                    Me.Lista118.Items.Add(" OR ")
                End If
                Me.Lista118.Items.Add("{VISTA_MAESTRO_DE_PERSONAS.FEC_NAC} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ") To CDateTime (" & A2 & ", " & M2 & ", " & D2 & ")")
            End If
            If CODIGO_REPORTE = "S0004" Or CODIGO_REPORTE = "S0005" Then
                If Me.Lista118.Items.Count <> 0 Then
                    Me.Lista118.Items.Add(" OR ")
                End If
                Me.Lista118.Items.Add("((DAY({VISTA_MAESTRO_DE_PERSONAS.FEC_NAC}) >= " & Me.DateTimePicker1.Value.Day & ") And (DAY({VISTA_MAESTRO_DE_PERSONAS.FEC_NAC}) <= " & Me.DateTimePicker2.Value.Day & ") And (MONTH({VISTA_MAESTRO_DE_PERSONAS.FEC_NAC}) >= " & Me.DateTimePicker1.Value.Month & ") And (MONTH({VISTA_MAESTRO_DE_PERSONAS.FEC_NAC}) <= " & Me.DateTimePicker2.Value.Month & ")) AND ({VISTA_MAESTRO_DE_CARGOS_SIT.ID_SITUACION} = 1)")
            End If
        End If
        If Me.RadioButton8.Checked = True Then         'FECHA DE SITUACION
            If Mid(Me.DateTimePicker1.Value, 1, 10) = Mid(Me.DateTimePicker2.Value, 1, 10) Then
                Me.Lista_Principal.Items.Add("Fecha De Situacion: " & Mid(Me.DateTimePicker1.Value, 1, 10))
            Else
                Me.Lista_Principal.Items.Add("Fecha De Situacion: Del " & Mid(Me.DateTimePicker1.Value, 1, 10) & " al " & Mid(Me.DateTimePicker2.Value, 1, 10))
            End If
            D1 = Mid(Me.DateTimePicker1.Value, 1, 2)
            M1 = Mid(Me.DateTimePicker1.Value, 4, 2)
            A1 = Mid(Me.DateTimePicker1.Value, 7, 4)
            D2 = Mid(Me.DateTimePicker2.Value, 1, 2)
            M2 = Mid(Me.DateTimePicker2.Value, 4, 2)
            A2 = Mid(Me.DateTimePicker2.Value, 7, 4)
            If Mid(Me.DateTimePicker1.Value, 1, 10) = Mid(Me.DateTimePicker2.Value, 1, 10) Then
                If Me.Lista121.Items.Count = 0 Then
                    Me.Lista121.Items.Add("{MAESTRO_DE_SITUACION_DE_PERSONAL.DIA} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ")")
                Else
                    Me.Lista121.Items.Add(" OR ")
                    Me.Lista121.Items.Add("{MAESTRO_DE_SITUACION_DE_PERSONAL.DIA} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ")")
                End If
            Else
                If Me.Lista121.Items.Count = 0 Then
                    Me.Lista121.Items.Add("{MAESTRO_DE_SITUACION_DE_PERSONAL.DIA} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ") To CDateTime (" & A2 & ", " & M2 & ", " & D2 & ")")
                Else
                    Me.Lista121.Items.Add(" OR ")
                    Me.Lista121.Items.Add("{MAESTRO_DE_SITUACION_DE_PERSONAL.DIA} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ") To CDateTime (" & A2 & ", " & M2 & ", " & D2 & ")")
                End If
            End If
        End If
        If Me.RadioButton9.Checked = True Then         'FECHA DE MARCAS
            If Mid(Me.DateTimePicker1.Value, 1, 10) = Mid(Me.DateTimePicker2.Value, 1, 10) Then
                Me.Lista_Principal.Items.Add("Fecha De Marcas: " & Mid(Me.DateTimePicker1.Value, 1, 10))
            Else
                Me.Lista_Principal.Items.Add("Fecha De Marcas: Del " & Mid(Me.DateTimePicker1.Value, 1, 10) & " al " & Mid(Me.DateTimePicker2.Value, 1, 10))
            End If
            D1 = Mid(Me.DateTimePicker1.Value, 1, 2)
            M1 = Mid(Me.DateTimePicker1.Value, 4, 2)
            A1 = Mid(Me.DateTimePicker1.Value, 7, 4)
            D2 = Mid(Me.DateTimePicker2.Value, 1, 2)
            M2 = Mid(Me.DateTimePicker2.Value, 4, 2)
            A2 = Mid(Me.DateTimePicker2.Value, 7, 4)
            If Mid(Me.DateTimePicker1.Value, 1, 10) = Mid(Me.DateTimePicker2.Value, 1, 10) Then
                If Me.Lista122.Items.Count = 0 Then
                    Me.Lista122.Items.Add("({MAESTRO_DE_MARCAS.HORA_MARCA} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ", 00, 00, 00) TO CDateTime (" & A1 & ", " & M1 & ", " & D1 & ", 23, 59, 00))")
                Else
                    Me.Lista122.Items.Add(" OR ")
                    Me.Lista122.Items.Add("({MAESTRO_DE_MARCAS.HORA_MARCA} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ", 00, 00, 00) TO CDateTime (" & A1 & ", " & M1 & ", " & D1 & ", 23, 59, 00))")
                End If
            Else
                If Me.Lista122.Items.Count = 0 Then
                    Me.Lista122.Items.Add("({MAESTRO_DE_MARCAS.HORA_MARCA} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ", 00, 00, 00) TO CDateTime (" & A2 & ", " & M2 & ", " & D2 & ", 23, 59, 00))")
                Else
                    Me.Lista122.Items.Add(" OR ")
                    Me.Lista122.Items.Add("({MAESTRO_DE_MARCAS.HORA_MARCA} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ", 00, 00, 00) TO CDateTime (" & A2 & ", " & M2 & ", " & D2 & ", 23, 59, 00))")
                End If
            End If
        End If
        If Me.RadioButton10.Checked = True Then         'FECHA DEL PROCESO DEL ASPIRANTE
            If Mid(Me.DateTimePicker1.Value, 1, 10) = Mid(Me.DateTimePicker2.Value, 1, 10) Then
                Me.Lista_Principal.Items.Add("Fecha De Proceso del Aspirante: " & Mid(Me.DateTimePicker1.Value, 1, 10))
            Else
                Me.Lista_Principal.Items.Add("Fecha De Proceso del Aspirante: Del " & Mid(Me.DateTimePicker1.Value, 1, 10) & " al " & Mid(Me.DateTimePicker2.Value, 1, 10))
            End If
            D1 = Mid(Me.DateTimePicker1.Value, 1, 2)
            M1 = Mid(Me.DateTimePicker1.Value, 4, 2)
            A1 = Mid(Me.DateTimePicker1.Value, 7, 4)
            D2 = Mid(Me.DateTimePicker2.Value, 1, 2)
            M2 = Mid(Me.DateTimePicker2.Value, 4, 2)
            A2 = Mid(Me.DateTimePicker2.Value, 7, 4)
            If Mid(Me.DateTimePicker1.Value, 1, 10) = Mid(Me.DateTimePicker2.Value, 1, 10) Then
                If Me.Lista123.Items.Count = 0 Then
                    Me.Lista123.Items.Add("{SOL_MAESTRO_DE_SOLICITANTES_PROCESOS.FECHA_P} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ")")
                Else
                    Me.Lista123.Items.Add(" OR ")
                    Me.Lista123.Items.Add("{SOL_MAESTRO_DE_SOLICITANTES_PROCESOS.FECHA_P} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ")")
                End If
            Else
                If Me.Lista123.Items.Count = 0 Then
                    Me.Lista123.Items.Add("{SOL_MAESTRO_DE_SOLICITANTES_PROCESOS.FECHA_P} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ") To CDateTime (" & A2 & ", " & M2 & ", " & D2 & ")")
                Else
                    Me.Lista123.Items.Add(" OR ")
                    Me.Lista123.Items.Add("{SOL_MAESTRO_DE_SOLICITANTES_PROCESOS.FECHA_P} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ") To CDateTime (" & A2 & ", " & M2 & ", " & D2 & ")")
                End If
            End If
        End If
        If Me.RadioButton11.Checked = True Then         'FECHA DE APLICACION DE MOVIMIENTOS
            If Mid(Me.DateTimePicker1.Value, 1, 10) = Mid(Me.DateTimePicker2.Value, 1, 10) Then
                Me.Lista_Principal.Items.Add("Fecha De Aplic. de Movimientos: " & Mid(Me.DateTimePicker1.Value, 1, 10))
            Else
                Me.Lista_Principal.Items.Add("Fecha De Aplic. de Movimientos: Del " & Mid(Me.DateTimePicker1.Value, 1, 10) & " al " & Mid(Me.DateTimePicker2.Value, 1, 10))
            End If
            D1 = Mid(Me.DateTimePicker1.Value, 1, 2)
            M1 = Mid(Me.DateTimePicker1.Value, 4, 2)
            A1 = Mid(Me.DateTimePicker1.Value, 7, 4)
            D2 = Mid(Me.DateTimePicker2.Value, 1, 2)
            M2 = Mid(Me.DateTimePicker2.Value, 4, 2)
            A2 = Mid(Me.DateTimePicker2.Value, 7, 4)
            If Mid(Me.DateTimePicker1.Value, 1, 10) = Mid(Me.DateTimePicker2.Value, 1, 10) Then
                If Me.Lista124.Items.Count = 0 Then
                    Me.Lista124.Items.Add("{HISTORICO_DE_MOVIMIENTO.FEC_APLICA} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ")")
                Else
                    Me.Lista124.Items.Add(" OR ")
                    Me.Lista124.Items.Add("{HISTORICO_DE_MOVIMIENTO.FEC_APLICA} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ")")
                End If
            Else
                If Me.Lista124.Items.Count = 0 Then
                    Me.Lista124.Items.Add("{HISTORICO_DE_MOVIMIENTO.FEC_APLICA} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ") To CDateTime (" & A2 & ", " & M2 & ", " & D2 & ")")
                Else
                    Me.Lista124.Items.Add(" OR ")
                    Me.Lista124.Items.Add("{HISTORICO_DE_MOVIMIENTO.FEC_APLICA} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ") To CDateTime (" & A2 & ", " & M2 & ", " & D2 & ")")
                End If
            End If
        End If
        If Me.RadioButton15.Checked = True Then         'FECHA DE EVENTUALIDADES
            If Mid(Me.DateTimePicker1.Value, 1, 10) = Mid(Me.DateTimePicker2.Value, 1, 10) Then
                Me.Lista_Principal.Items.Add("Fecha de Eventualidad: " & Mid(Me.DateTimePicker1.Value, 1, 10))
            Else
                Me.Lista_Principal.Items.Add("Fecha de Eventualidad: Del " & Mid(Me.DateTimePicker1.Value, 1, 10) & " al " & Mid(Me.DateTimePicker2.Value, 1, 10))
            End If
            D1 = Mid(Me.DateTimePicker1.Value, 1, 2)
            M1 = Mid(Me.DateTimePicker1.Value, 4, 2)
            A1 = Mid(Me.DateTimePicker1.Value, 7, 4)
            D2 = Mid(Me.DateTimePicker2.Value, 1, 2)
            M2 = Mid(Me.DateTimePicker2.Value, 4, 2)
            A2 = Mid(Me.DateTimePicker2.Value, 7, 4)
            If Mid(Me.DateTimePicker1.Value, 1, 10) = Mid(Me.DateTimePicker2.Value, 1, 10) Then
                If Me.Lista118.Items.Count = 0 Then
                    Me.Lista118.Items.Add("{MAESTRO_DE_EVENTUALIDADES.FEC_EVE} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ")")
                Else
                    Me.Lista118.Items.Add(" OR ")
                    Me.Lista118.Items.Add("{MAESTRO_DE_EVENTUALIDADES.FEC_EVE} = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ")")
                End If
            Else
                If Me.Lista118.Items.Count = 0 Then
                    Me.Lista118.Items.Add("{MAESTRO_DE_EVENTUALIDADES.FEC_EVE}  = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ") To CDateTime (" & A2 & ", " & M2 & ", " & D2 & ")")
                Else
                    Me.Lista118.Items.Add(" OR ")
                    Me.Lista118.Items.Add("{MAESTRO_DE_EVENTUALIDADES.FEC_EVE}  = CDateTime (" & A1 & ", " & M1 & ", " & D1 & ") To CDateTime (" & A2 & ", " & M2 & ", " & D2 & ")")
                End If
            End If
        End If
        If Me.RadioButton16.Checked = True Then         'FECHA SE SUBSIDIOS
            If Mid(Me.DateTimePicker1.Value, 1, 10) = Mid(Me.DateTimePicker2.Value, 1, 10) Then
                Me.Lista_Principal.Items.Add("Fecha de Subsidio: " & Mid(Me.DateTimePicker1.Value, 1, 10))
            Else
                Me.Lista_Principal.Items.Add("Fecha de Subsidio: Del " & Mid(Me.DateTimePicker1.Value, 1, 10) & " al " & Mid(Me.DateTimePicker2.Value, 1, 10))
            End If
            D1 = Mid(Me.DateTimePicker1.Value, 1, 2)
            M1 = Mid(Me.DateTimePicker1.Value, 4, 2)
            A1 = Mid(Me.DateTimePicker1.Value, 7, 4)
            D2 = Mid(Me.DateTimePicker2.Value, 1, 2)
            M2 = Mid(Me.DateTimePicker2.Value, 4, 2)
            A2 = Mid(Me.DateTimePicker2.Value, 7, 4)
            If CODIGO_REPORTE = "S0005" Or CODIGO_REPORTE = "S0006" Then
                If Me.Lista140.Items.Count <> 0 Then
                    Me.Lista140.Items.Add(" OR ")
                End If
                Me.Lista140.Items.Add("({MAESTRO_DE_SUBSIDIOS.FECHA_I} >= CDateTime (" & A1 & ", " & M1 & ", " & D1 & ", 00, 00, 00) AND {MAESTRO_DE_SUBSIDIOS.FECHA_F} <= CDateTime (" & A2 & ", " & M2 & ", " & D2 & ", 00, 00, 00))")
            End If
        End If
    End Sub
    'ASPIRANTES PROCESOS
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        If Me.ComboBox17.SelectedValue <> 0 Then
            Me.Lista_Principal.Items.Add("Proceso del Aspirante: " & Me.ComboBox17.Text)
            If Me.Lista125.Items.Count = 0 Then
                Me.Lista125.Items.Add("({SOL_MAESTRO_DE_SOLICITANTES_PROCESOS.ID_PROC} = " & Me.ComboBox17.SelectedValue & ")")
            Else
                Me.Lista125.Items.Add(" OR ")
                Me.Lista125.Items.Add("({SOL_MAESTRO_DE_SOLICITANTES_PROCESOS.ID_PROC} = " & Me.ComboBox17.SelectedValue & ")")
            End If
        Else
            MsgBox("La Seleccion para el Proceso del Aspirante no es válida", vbInformation, "Mensaje del Sistema")
            Me.ComboBox17.Focus()
            Exit Sub
        End If
    End Sub
    'ASPIRANTES RESULTADOS
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        If Me.ComboBox18.SelectedValue <> 0 Then
            Me.Lista_Principal.Items.Add("Resultado del Proceso del Aspirante: " & Me.ComboBox18.Text)
            If Me.Lista126.Items.Count = 0 Then
                Me.Lista126.Items.Add("({SOL_MAESTRO_DE_SOLICITANTES_PROCESOS.ID_R_SOL} = " & Me.ComboBox18.SelectedValue & ")")
            Else
                Me.Lista126.Items.Add(" OR ")
                Me.Lista126.Items.Add("({SOL_MAESTRO_DE_SOLICITANTES_PROCESOS.ID_R_SOL} = " & Me.ComboBox18.SelectedValue & ")")
            End If
        Else
            MsgBox("La Seleccion para el Resultado del Proceso del Aspirante no es válida", vbInformation, "Mensaje del Sistema")
            Me.ComboBox18.Focus()
            Exit Sub
        End If
    End Sub
    'FUNCIONALIDAD
    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        If Me.ComboBox19.SelectedValue <> 0 Then
            Me.Lista_Principal.Items.Add("Funcionalidad: " & Me.ComboBox19.Text)
            If Me.Lista127.Items.Count = 0 Then
                If CODIGO_REPORTE = "E0002" Then
                    Me.Lista127.Items.Add("({VISTA_MAESTRO_DE_PERSONAS.ID_F} = " & Me.ComboBox19.SelectedValue & ")")
                Else
                    Me.Lista127.Items.Add("({MAESTRO_DE_PERSONAS.ID_F} = " & Me.ComboBox19.SelectedValue & ")")
                End If
            Else
                Me.Lista127.Items.Add(" OR ")
                If CODIGO_REPORTE = "E0002" Then
                    Me.Lista127.Items.Add("({VISTA_MAESTRO_DE_PERSONAS.ID_F} = " & Me.ComboBox19.SelectedValue & ")")
                Else
                    Me.Lista127.Items.Add("({MAESTRO_DE_PERSONAS.ID_F} = " & Me.ComboBox19.SelectedValue & ")")
                End If
            End If
        Else
            MsgBox("La Seleccion para la Funcionalidad no es válida", vbInformation, "Mensaje del Sistema")
            Me.ComboBox19.Focus()
            Exit Sub
        End If
    End Sub
    'CAUSA REAL DE EVENTUALIDAD
    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        If Me.ComboBox20.SelectedValue <> 0 Then
            Me.Lista_Principal.Items.Add("Causa Real Eventualidad: " & Me.ComboBox20.Text)
            If Me.Lista128.Items.Count = 0 Then
                Me.Lista128.Items.Add("({MAESTRO_DE_EVENTUALIDADES.ID_CR_EVENTOS} = " & Me.ComboBox20.SelectedValue & ")")
            Else
                Me.Lista128.Items.Add(" OR ")
                Me.Lista128.Items.Add("({MAESTRO_DE_EVENTUALIDADES.ID_CR_EVENTOS} = " & Me.ComboBox20.SelectedValue & ")")
            End If
        Else
            MsgBox("La Seleccion para la Causa Real Eventualidad no es válida", vbInformation, "Mensaje del Sistema")
            Me.ComboBox20.Focus()
            Exit Sub
        End If
    End Sub
    'TIPO DE EVENTUALIDAD
    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        If Me.ComboBox21.SelectedValue <> 0 Then
            Me.Lista_Principal.Items.Add("Tipo de Eventualidad: " & Me.ComboBox21.Text)
            If Me.Lista129.Items.Count = 0 Then
                Me.Lista129.Items.Add("({MAESTRO_DE_EVENTUALIDADES.ID_E} = " & Me.ComboBox21.SelectedValue & ")")
            Else
                Me.Lista129.Items.Add(" OR ")
                Me.Lista129.Items.Add("({MAESTRO_DE_EVENTUALIDADES.ID_E} = " & Me.ComboBox21.SelectedValue & ")")
            End If
        Else
            MsgBox("La Seleccion para el Tipo de Eventualidad no es válida", vbInformation, "Mensaje del Sistema")
            Me.ComboBox21.Focus()
            Exit Sub
        End If
    End Sub
    'TIPO DE ESTUDIO
    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        If Me.ComboBox22.SelectedValue <> 0 Then
            Me.Lista_Principal.Items.Add("Tipo de Estudio: " & Me.ComboBox22.Text)
            If Me.Lista130.Items.Count = 0 Then
                Me.Lista130.Items.Add("({MAESTRO_DE_ESTUDIOS.ID_T_ESTUDIO} = " & Me.ComboBox22.SelectedValue & ")")
            Else
                Me.Lista130.Items.Add(" OR ")
                Me.Lista130.Items.Add("({MAESTRO_DE_ESTUDIOS.ID_T_ESTUDIO} = " & Me.ComboBox22.SelectedValue & ")")
            End If
        Else
            MsgBox("La Seleccion para el Tipo de Estudio no es válida", vbInformation, "Mensaje del Sistema")
            Me.ComboBox22.Focus()
            Exit Sub
        End If
    End Sub
    'NIVEL DE COMPETENCIA
    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        If Me.ComboBox23.SelectedValue <> 0 Then
            Me.Lista_Principal.Items.Add("Nivel de Competencia: " & Me.ComboBox23.Text)
            If Me.Lista131.Items.Count = 0 Then
                Me.Lista131.Items.Add("({MAESTRO_DE_PERSONAS.ID_N_COMPETENCIA} = " & Me.ComboBox23.SelectedValue & ")")
            Else
                Me.Lista131.Items.Add(" OR ")
                Me.Lista131.Items.Add("({MAESTRO_DE_ESTUDIOS.ID_N_COMPETENCIA} = " & Me.ComboBox23.SelectedValue & ")")
            End If
        Else
            MsgBox("La Seleccion para el Nivel de Competencia no es válida", vbInformation, "Mensaje del Sistema")
            Me.ComboBox23.Focus()
            Exit Sub
        End If
    End Sub
    'RELOJ DE MARCACION
    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        If Me.ComboBox24.Text <> "" Then
            Me.Lista_Principal.Items.Add("Reloj de Marcacion: " & Me.ComboBox24.Text)
            If Me.Lista132.Items.Count = 0 Then
                Me.Lista132.Items.Add("({MAESTRO_DE_MARCAS.IDRELOJ} = '" & Me.ComboBox24.SelectedValue & "')")
            Else
                Me.Lista132.Items.Add(" OR ")
                Me.Lista132.Items.Add("({MAESTRO_DE_MARCAS.IDRELOJ} = '" & Me.ComboBox24.SelectedValue & "')")
            End If
        Else
            MsgBox("La Seleccion para el Reloj de Marcacion no es válida", vbInformation, "Mensaje del Sistema")
            Me.ComboBox24.Focus()
            Exit Sub
        End If
    End Sub
    'TIPO DE MARCA
    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        If Me.ComboBox25.Text <> "" Then
            Me.Lista_Principal.Items.Add("Tipo de Marca: " & Me.ComboBox25.Text)
            If Me.Lista133.Items.Count = 0 Then
                Me.Lista133.Items.Add("({MAESTRO_DE_MARCAS.TIPO_MARCA} = '" & Me.ComboBox25.Text & "')")
            Else
                Me.Lista133.Items.Add(" OR ")
                Me.Lista133.Items.Add("({MAESTRO_DE_MARCAS.TIPO_MARCA} = '" & Me.ComboBox25.Text & "')")
            End If
        Else
            MsgBox("La Seleccion para el Tipo de Marca no es válida", vbInformation, "Mensaje del Sistema")
            Me.ComboBox25.Focus()
            Exit Sub
        End If
    End Sub
    'CLASIFICACION DEL EMPLEADO
    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        If Me.ComboBox26.SelectedValue <> 0 Then
            Me.Lista_Principal.Items.Add("Clasificacion del Empleado: " & Me.ComboBox26.Text)
            If Me.Lista134.Items.Count = 0 Then
                Me.Lista134.Items.Add("({MAESTRO_DE_CLASIFICACION_DE_PERSONAL.ID_C_PE} = " & Me.ComboBox26.SelectedValue & ")")
            Else
                Me.Lista134.Items.Add(" OR ")
                Me.Lista134.Items.Add("({MAESTRO_DE_CLASIFICACION_DE_PERSONAL.ID_C_PE} = " & Me.ComboBox26.SelectedValue & ")")
            End If
        Else
            MsgBox("La Seleccion para la Clasificacion del Empleado no es válida", vbInformation, "Mensaje del Sistema")
            Me.ComboBox26.Focus()
            Exit Sub
        End If
    End Sub
    'CODIGO DEL EMPLEADO
    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        cEMPLEADO = CInt(Val(TextBox1.Text))
        TextBox1.Text = Format(CInt(cEMPLEADO), "0000000000")
        If TextBox1.Text = "0000000000" Then
            MsgBox("Debe digitar un Codigo de Empleado válido", vbInformation, "Mensaje del Sistema")
            Me.TextBox1.Focus()
            Me.TextBox1.SelectAll()
            Exit Sub
        End If
        Call BUSCAR_ID_EMPLEADO()
        If EXISTE_ID = True Then
            Me.Lista_Principal.Items.Add("Código del Empleado: " & Me.TextBox1.Text & " " & vAN)
            If Me.Lista135.Items.Count = 0 Then
                If CODIGO_REPORTE = "E0002" Then
                    Me.Lista135.Items.Add("({VISTA_MAESTRO_DE_PERSONAS.CODIGO} = '" & Me.TextBox1.Text & "')")
                Else
                    If CODIGO_REPORTE = "E0005" Then
                        Me.Lista135.Items.Add("({VISTA_CONSULTA_MAESTRO_DE_CURVA_Y_TALLA.CODIGO} = '" & Me.TextBox1.Text & "')")
                    Else
                        If CODIGO_REPORTE = "S0001" Then
                            Me.Lista135.Items.Add("({MAESTRO_DE_MARCAS.ID_EMPLEADO} = '" & Me.TextBox1.Text & "')")
                        Else
                            Me.Lista135.Items.Add("({MAESTRO_DE_PERSONAS.CODIGO} = '" & Me.TextBox1.Text & "')")
                        End If
                    End If
                End If
            Else
                Me.Lista135.Items.Add(" OR ")
                If CODIGO_REPORTE = "E0002" Then
                    Me.Lista135.Items.Add("({VISTA_MAESTRO_DE_PERSONAS.CODIGO} = '" & Me.TextBox1.Text & "')")
                Else
                    If CODIGO_REPORTE = "E0005" Then
                        Me.Lista135.Items.Add("({VISTA_CONSULTA_MAESTRO_DE_CURVA_Y_TALLA.CODIGO} = '" & Me.TextBox1.Text & "')")
                    Else
                        If CODIGO_REPORTE = "S0001" Then
                            Me.Lista135.Items.Add("({MAESTRO_DE_MARCAS.ID_EMPLEADO} = '" & Me.TextBox1.Text & "')")
                        Else
                            Me.Lista135.Items.Add("({MAESTRO_DE_PERSONAS.CODIGO} = '" & Me.TextBox1.Text & "')")
                        End If
                    End If
                End If
            End If
            Me.TextBox1.Text = ""
            Me.TextBox1.SelectAll()
            Me.TextBox1.Focus()
        Else
            If CODIGO_REPORTE = "S0001" Then
                Me.Lista_Principal.Items.Add("Código del Empleado: " & Me.TextBox1.Text & " " & vAN)
                If Me.Lista135.Items.Count = 0 Then
                    If CODIGO_REPORTE = "S0001" Then
                        Me.Lista135.Items.Add("({MAESTRO_DE_MARCAS.ID_EMPLEADO} = '" & Me.TextBox1.Text & "')")
                    End If
                End If
            Else
                MsgBox("La Seleccion para el Código del Empleado no es válida", vbInformation, "Mensaje del Sistema")
                Me.TextBox1.SelectAll()
                Me.TextBox1.Focus()
                Exit Sub
            End If
        End If
    End Sub
    Dim EXISTE_ID As Boolean
    Dim vID_M_P As Integer
    Dim vAN As String
    Dim vCOD As String
    Dim vACT As Boolean
    Private Sub BUSCAR_ID_EMPLEADO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT ID_M_P, CODIGO, APELLIDOS, NOMBRES, ID_ESTADO_P FROM [dbo].[MAESTRO DE PERSONAS] WHERE CODIGO = '" & Me.TextBox1.Text & "'", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                vACT = MiDataTable.Rows(0).Item(4).ToString
                vID_M_P = MiDataTable.Rows(0).Item(0).ToString
                vCOD = MiDataTable.Rows(0).Item(1).ToString
                vAN = Trim(MiDataTable.Rows(0).Item(2).ToString) & " " & Trim(MiDataTable.Rows(0).Item(3).ToString)
                If vACT = True Then
                    EXISTE_ID = True

                Else
                    '    id_EMPLEADO = 0
                    '    vCOD = ""
                    '    vAN = ""
                    EXISTE_ID = False
                End If
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    'VALIDAR QUE SOLO INGRESEN NUMEROS (codigo del empleado)
    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Char.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsSymbol(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsSeparator(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsWhiteSpace(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
        Me.TextBox1.Text = Trim(Replace(Me.TextBox1.Text, "  ", " "))
        TextBox1.Select(TextBox1.Text.Length, 0)
    End Sub
    'ANTIGUEDAD
    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        If Me.RadioButton12.Checked = False And Me.RadioButton13.Checked = False And Me.RadioButton14.Checked = False Then
            MsgBox("Debe seleccionar el Tipo de Edad o Antiguedad Correctamente", vbInformation, "Mensaje del Sistema")
            Me.RadioButton12.Focus()
            Exit Sub
        End If
        If Me.TextBox2.Text = "" Or Me.TextBox3.Text = "" Then
            MsgBox("Debe digitar los Valores Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox2.Focus()
            Exit Sub
        End If

        If Me.RadioButton12.Checked = True Then         'EDAD
            If Trim(Me.TextBox2.Text) = Trim(Me.TextBox3.Text) Then
                Me.Lista_Principal.Items.Add("Edad: " & Me.TextBox2.Text & " año(s)")
            Else
                Me.Lista_Principal.Items.Add("Edad: De " & Me.TextBox2.Text & " a " & Me.TextBox3.Text & " año(s)")
            End If

            If Me.Lista136.Items.Count <> 0 Then
                Me.Lista136.Items.Add(" OR ")
            End If
            Me.Lista136.Items.Add("({VISTA_MAESTRO_DE_PERSONAS.EDAD} >= " & Me.TextBox2.Text & " AND {VISTA_MAESTRO_DE_PERSONAS.EDAD} <= " & Me.TextBox3.Text & ")")
        End If

        If Me.RadioButton13.Checked = True Then         'ANTIGUEDAD PAME
            If Trim(Me.TextBox2.Text) = Trim(Me.TextBox3.Text) Then
                Me.Lista_Principal.Items.Add("Antiguedad PAME: " & Me.TextBox2.Text & " año(s)")
            Else
                Me.Lista_Principal.Items.Add("Antiguedad PAME: De " & Me.TextBox2.Text & " a " & Me.TextBox3.Text & " año(s)")
            End If
            If Trim(Me.TextBox2.Text) = Trim(Me.TextBox3.Text) Then
                If Me.Lista136.Items.Count = 0 Then
                    Me.Lista136.Items.Add("({VISTA_MAESTRO_DE_PERSONAS.ANTIGUEDAD_PAME} >= " & Me.TextBox2.Text & ")")
                Else
                    Me.Lista136.Items.Add(" OR ")
                    Me.Lista136.Items.Add("({VISTA_MAESTRO_DE_PERSONAS.ANTIGUEDAD_PAME} = " & Me.TextBox2.Text & ")")
                End If
            Else
                If Me.Lista136.Items.Count = 0 Then
                    Me.Lista136.Items.Add("({VISTA_MAESTRO_DE_PERSONAS.ANTIGUEDAD_PAME} >= " & Me.TextBox2.Text & " AND {VISTA_MAESTRO_DE_PERSONAS.ANTIGUEDAD_PAME} <= " & Me.TextBox3.Text & ")")
                Else
                    Me.Lista136.Items.Add(" OR ")
                    Me.Lista136.Items.Add("({VISTA_MAESTRO_DE_PERSONAS.ANTIGUEDAD_PAME} >= " & Me.TextBox2.Text & " AND {VISTA_MAESTRO_DE_PERSONAS.ANTIGUEDAD_PAME} <= " & Me.TextBox3.Text & ")")
                End If
            End If
        End If
        If Me.RadioButton14.Checked = True Then         'ANTIGUEDAD EN
            If Trim(Me.TextBox2.Text) = Trim(Me.TextBox3.Text) Then
                Me.Lista_Principal.Items.Add("Antiguedad EN: " & Me.TextBox2.Text & " año(s)")
            Else
                Me.Lista_Principal.Items.Add("Antiguedad EN: De " & Me.TextBox2.Text & " a " & Me.TextBox3.Text & " año(s)")
            End If
            If Trim(Me.TextBox2.Text) = Trim(Me.TextBox3.Text) Then
                If Me.Lista136.Items.Count = 0 Then
                    Me.Lista136.Items.Add("({VISTA_MAESTRO_DE_PERSONAS.ANTIGUEDAD_EN} >= " & Me.TextBox2.Text & ")")
                Else
                    Me.Lista136.Items.Add(" OR ")
                    Me.Lista136.Items.Add("({VISTA_MAESTRO_DE_PERSONAS.ANTIGUEDAD_EN} = " & Me.TextBox2.Text & ")")
                End If
            Else
                If Me.Lista136.Items.Count = 0 Then
                    Me.Lista136.Items.Add("({VISTA_MAESTRO_DE_PERSONAS.ANTIGUEDAD_EN} >= " & Me.TextBox2.Text & " AND {VISTA_MAESTRO_DE_PERSONAS.ANTIGUEDAD_EN} <= " & Me.TextBox3.Text & ")")
                Else
                    Me.Lista136.Items.Add(" OR ")
                    Me.Lista136.Items.Add("({VISTA_MAESTRO_DE_PERSONAS.ANTIGUEDAD_EN} >= " & Me.TextBox2.Text & " AND {VISTA_MAESTRO_DE_PERSONAS.ANTIGUEDAD_EN} <= " & Me.TextBox3.Text & ")")
                End If
            End If
        End If
    End Sub
    'VALIDAR QUE SOLO INGRESEN NUMEROS (antiguedad)
    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Char.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsSymbol(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsSeparator(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsWhiteSpace(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
        Me.TextBox2.Text = Trim(Replace(Me.TextBox2.Text, "  ", " "))
        TextBox2.Select(TextBox2.Text.Length, 0)
    End Sub
    'VALIDAR QUE SOLO INGRESEN NUMEROS (antiguedad)
    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        If Char.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsSymbol(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsSeparator(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsWhiteSpace(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
        Me.TextBox3.Text = Trim(Replace(Me.TextBox3.Text, "  ", " "))
        TextBox3.Select(TextBox3.Text.Length, 0)
    End Sub
    'CONTRATOS
    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        If Me.TextBox4.Text = "" Then
            MsgBox("Debe digitar los dias Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox4.Focus()
            Exit Sub
        End If
        If Me.TextBox4.Text > 0 Then
            Dim FECHA_CORTE As Date
            Call Clases.BUSCAR_FECHA_Y_HORA()
            FECHA_CORTE = FECHA_SERVIDOR.AddDays(CInt(Me.TextBox4.Text))
            Dim Df1 As String
            Df1 = Mid(FECHA_SERVIDOR, 1, 2)
            Dim Mf1 As String
            Mf1 = Mid(FECHA_SERVIDOR, 4, 2)
            Dim Af1 As String
            Af1 = Mid(FECHA_SERVIDOR, 7, 4)
            Dim Df2 As String
            Df2 = Mid(FECHA_CORTE, 1, 2)
            Dim Mf2 As String
            Mf2 = Mid(FECHA_CORTE, 4, 2)
            Dim Af2 As String
            Af2 = Mid(FECHA_CORTE, 7, 4)
            Me.Lista_Principal.Items.Add("Contratos a Vencer en: " & Me.TextBox4.Text & " dia(s)")
            If Me.Lista137.Items.Count = 0 Then
                Me.Lista137.Items.Add("{MAESTRO_DE_CONTRATOS_Y_EVALUACIONES.FECHA_F} = CDateTime (" & Af1 & ", " & Mf1 & ", " & Df1 & ") TO CDateTime (" & Af2 & ", " & Mf2 & ", " & Df2 & ")")
            Else
                Me.Lista137.Items.Add(" OR ")
                Me.Lista137.Items.Add("{MAESTRO_DE_CONTRATOS_Y_EVALUACIONES.FECHA_F} = CDateTime (" & Af1 & ", " & Mf1 & ", " & Df1 & ") TO CDateTime (" & Af2 & ", " & Mf2 & ", " & Df2 & ")")
            End If
        Else
            MsgBox("La Seleccion para la Cantidad de Dias no es válida", vbInformation, "Mensaje del Sistema")
            Me.TextBox4.SelectAll()
            Me.TextBox4.Focus()
            Exit Sub
        End If
    End Sub
    Private Sub Button100_Click(sender As Object, e As EventArgs) Handles Button100.Click
        Call VERIFICAR_ACCESOS("300") : If HAY_ACCESO = False Then : Exit Sub : End If
        If NOMBRE_REPORTE = "" Then
            MsgBox("Debe seleccionar el Reporte que desea visualizar", vbInformation, "Mensaje del Sistema")
            Exit Sub
        End If
        '***********************
        If NOMBRE_REPORTE = "O0001 - Catalogo de Organización" Then : Call VERIFICAR_ACCESOS("318") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "O0002 - Lista de Organica (MALIB) por Niveles Organizacionales" Then : Call VERIFICAR_ACCESOS("319") : If HAY_ACCESO = False Then : Exit Sub : End If : End If

        If NOMBRE_REPORTE = "O0003 - Parte de Organica por Niveles Organizacionales" Then : Call VERIFICAR_ACCESOS("320") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "O0004 - Parte Especialistas por Niveles Organizacionales" Then : Call VERIFICAR_ACCESOS("321") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "O0005 - Lista Especialistas por Niveles Organizacionales" Then : Call VERIFICAR_ACCESOS("322") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "O0006 - Parte Sub Especialistas por Niveles Organizacionales" Then : Call VERIFICAR_ACCESOS("323") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "O0007 - Lista Sub Especialistas por Niveles Organizacionales" Then : Call VERIFICAR_ACCESOS("324") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "O0008 - Parte de Cargos por Niveles Organizacionales" Then : Call VERIFICAR_ACCESOS("325") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "O0009 - Parte de Funcionalidades de Cargos por Niveles Organizacionales" Then : Call VERIFICAR_ACCESOS("326") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "O0010 - Parte de Organica (MALIB) por Establecimiento" Then : Call VERIFICAR_ACCESOS("327") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "O0011 - Lista de Organica (MALIB) por Establecimiento" Then : Call VERIFICAR_ACCESOS("328") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "M0001 - Resumen de Altas y Bajas por Mes" Then : Call VERIFICAR_ACCESOS("329") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "M0002 - Resumen de Altas y Bajas por Niveles Organizacionales" Then : Call VERIFICAR_ACCESOS("330") : If HAY_ACCESO = False Then : Exit Sub : End If : End If

        If NOMBRE_REPORTE = "M0003 - Resumen de Altas y Bajas por Cargos" Then : Call VERIFICAR_ACCESOS("331") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "M0004 - Resumen de Movimientos por Niveles Organizacionales y Rango de Fechas" Then : Call VERIFICAR_ACCESOS("332") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "M0005 - Resumen de Movimientos por Cargos y Rango de Fechas" Then : Call VERIFICAR_ACCESOS("333") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "M0006 - Resumen de Movimientos por Causales Reales y Cargos" Then : Call VERIFICAR_ACCESOS("334") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "M0007 - Resumen de Movimientos por Causales Legales y Cargos" Then : Call VERIFICAR_ACCESOS("335") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "M0008 - Resumen de Movimientos por Causales Reales y Niveles Organizacionales" Then : Call VERIFICAR_ACCESOS("336") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "M0009 - Resumen de Movimientos por Causales Legales y Niveles Organizacionales" Then : Call VERIFICAR_ACCESOS("337") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "M0010 - Lista de Movimientos con sus Detalles" Then : Call VERIFICAR_ACCESOS("338") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "M0011 - Resumen de Movimientos por Causales Reales y Rango de Fechas" Then : Call VERIFICAR_ACCESOS("375") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "A0001 - Lista de Aspirantes" Then : Call VERIFICAR_ACCESOS("339") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "E0001 - Expediente de Empleado (ACTIVOS - INACTIVOS)" Then : Call VERIFICAR_ACCESOS("340") : If HAY_ACCESO = False Then : Exit Sub : End If : End If

        If NOMBRE_REPORTE = "E0002 - Empleados por Antiguedad (Edad, Antiguedad PAME, Antiguedad EN)" Then : Call VERIFICAR_ACCESOS("341") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "E0003 - Lista de Empleados por Clasificacion" Then : Call VERIFICAR_ACCESOS("342") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "E0004 - Historico de Condecoraciones por Empleado" Then : Call VERIFICAR_ACCESOS("343") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "E0005 - Curva y Talla de Empleados por Niveles Organizacionales" Then : Call VERIFICAR_ACCESOS("344") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "E0006 - Lista de Empleados por Funcionalidad" Then : Call VERIFICAR_ACCESOS("345") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "E0007 - Contratos y Evaluaciones por Empleado" Then : Call VERIFICAR_ACCESOS("346") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "E0008 - Eventualidades" Then : Call VERIFICAR_ACCESOS("347") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "E0009 - Constancia de Personal Activo" Then : Call VERIFICAR_ACCESOS("348") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "E0010 - Constancia de Personal Inactivo" Then : Call VERIFICAR_ACCESOS("349") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "E0011 - Constancia Salarial" Then : Call VERIFICAR_ACCESOS("350") : If HAY_ACCESO = False Then : Exit Sub : End If : End If

        If NOMBRE_REPORTE = "E0012 - Solvencia" Then : Call VERIFICAR_ACCESOS("351") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "E0013 - Aplicacion del reglamento disciplinario" Then : Call VERIFICAR_ACCESOS("352") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "E0014 - Documentos Digitalizados por Colaborador" Then : Call VERIFICAR_ACCESOS("353") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "E0015 - Lista de Colaboradores con su Foto" Then : Call VERIFICAR_ACCESOS("354") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "C0001 - Lista de Medicos Acreditados" Then : Call VERIFICAR_ACCESOS("355") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "B0001 - Lista de Becados" Then : Call VERIFICAR_ACCESOS("356") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "B0002 - Constancia de Becados" Then : Call VERIFICAR_ACCESOS("357") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "S0001 - Marcas por Empleado (ACTIVOS - INACTIVOS)" Then : Call VERIFICAR_ACCESOS("358") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "S0002 - Marcas por Niveles Organizacionales" Then : Call VERIFICAR_ACCESOS("359") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "S0003 - Formato de Solicitud de Vacaciones" Then : Call VERIFICAR_ACCESOS("360") : If HAY_ACCESO = False Then : Exit Sub : End If : End If

        If NOMBRE_REPORTE = "S0004 - Cumpleañeros" Then : Call VERIFICAR_ACCESOS("361") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "S0005 - Subsidios por Niveles Organizacionales" Then : Call VERIFICAR_ACCESOS("362") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "S0006 - Detalles de Subsidios por Niveles Organizacionales" Then : Call VERIFICAR_ACCESOS("363") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "S0007 - Situaciones del Personal por Rango de Fechas" Then : Call VERIFICAR_ACCESOS("364") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "S0008 - Estado Vacacional por Rango de Fechas (Detalles)" Then : Call VERIFICAR_ACCESOS("365") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "S0009 - Estado Vacacional por Rango de Fechas (Niveles Organizacionales)" Then : Call VERIFICAR_ACCESOS("366") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "S0010 - Indicadores de < Cantidad de Empleados por Tipo de Situacion A >" Then : Call VERIFICAR_ACCESOS("367") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "S0011 - Indicadores de < Cantidad de Empleados por Tipo de Situacion B >" Then : Call VERIFICAR_ACCESOS("368") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "S0012 - Indicadores de < Cantidad de Dias por Empleado y Tipo de Situacion >" Then : Call VERIFICAR_ACCESOS("369") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "S0013 - Indicadores de < Lista de Empleados por Tipo de Situacion >" Then : Call VERIFICAR_ACCESOS("370") : If HAY_ACCESO = False Then : Exit Sub : End If : End If

        If NOMBRE_REPORTE = "S0014 - Colilla de Saldo Vacacional" Then : Call VERIFICAR_ACCESOS("371") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "D0001 - Dosimetria" Then : Call VERIFICAR_ACCESOS("372") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "D0002 - Lista de Colaboradores con Inmunizaciones" Then : Call VERIFICAR_ACCESOS("373") : If HAY_ACCESO = False Then : Exit Sub : End If : End If
        If NOMBRE_REPORTE = "D0003 - Lista de Colaboradores con Inmunizaciones 2" Then : Call VERIFICAR_ACCESOS("374") : If HAY_ACCESO = False Then : Exit Sub : End If : End If

        '***********************
        Call ARMAR_CADENA_FILTRO_SELECCION_REPORTE()
        Call ARMAR_CADENA_GENERAL()
        SELECCION = CADENA_LISTAS
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        'ORGANICA
        If NOMBRE_REPORTE = "O0001 - Catalogo de Organización" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCO100.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCO100x.rpt"
            PARAMETRO = 100
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "O0002 - Lista de Organica (MALIB) por Niveles Organizacionales" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCO101.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCO101x.rpt"
            PARAMETRO = 101
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "O0003 - Parte de Organica por Niveles Organizacionales" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCO102.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCO102x.rpt"
            PARAMETRO = 102
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "O0004 - Parte Especialistas por Niveles Organizacionales" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCO103.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCO103x.rpt"
            PARAMETRO = 103
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "O0005 - Lista Especialistas por Niveles Organizacionales" Then
            '''''''''''''''''''''''''''''''''''' AGREGAR SELECCION OBLIGATORIA, PARAMETRO ESPECIALISTA = TRUE
            If CADENA_LISTAS = "" Then
                CADENA_LISTAS = "({MAESTRO_DE_PERSONAS.ESPECIALISTA} = TRUE)"
            Else
                CADENA_LISTAS = CADENA_LISTAS & " AND ({MAESTRO_DE_PERSONAS.ESPECIALISTA} = TRUE)"
            End If
            SELECCION = CADENA_LISTAS
            ''''''''''''''''''''''''''''''''''''
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCO104.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCO104x.rpt"
            PARAMETRO = 104
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "O0006 - Parte Sub Especialistas por Niveles Organizacionales" Then
            '''''''''''''''''''''''''''''''''''' AGREGAR SELECCION OBLIGATORIA, PARAMETRO SUBESPECIALISTA = TRUE
            If CADENA_LISTAS = "" Then
                CADENA_LISTAS = "({MAESTRO_DE_PERSONAS.SUBESPECIALISTA} = TRUE)"
            Else
                CADENA_LISTAS = CADENA_LISTAS & " AND ({MAESTRO_DE_PERSONAS.SUBESPECIALISTA} = TRUE)"
            End If
            SELECCION = CADENA_LISTAS
            ''''''''''''''''''''''''''''''''''''
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCO105.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCO105x.rpt"
            PARAMETRO = 105
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "O0007 - Lista Sub Especialistas por Niveles Organizacionales" Then
            '''''''''''''''''''''''''''''''''''' AGREGAR SELECCION OBLIGATORIA, PARAMETRO SUBESPECIALISTA = TRUE
            If CADENA_LISTAS = "" Then
                CADENA_LISTAS = "({MAESTRO_DE_PERSONAS.SUBESPECIALISTA} = TRUE)"
            Else
                CADENA_LISTAS = CADENA_LISTAS & " AND ({MAESTRO_DE_PERSONAS.SUBESPECIALISTA} = TRUE)"
            End If
            SELECCION = CADENA_LISTAS
            ''''''''''''''''''''''''''''''''''''
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCO106.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCO106x.rpt"
            PARAMETRO = 106
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "O0008 - Parte de Cargos por Niveles Organizacionales" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCO107.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCO107x.rpt"
            PARAMETRO = 107
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "O0009 - Parte de Funcionalidades de Cargos por Niveles Organizacionales" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCO108.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCO108x.rpt"
            PARAMETRO = 108
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "O0010 - Parte de Organica (MALIB) por Establecimiento" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCO109.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCO109x.rpt"
            PARAMETRO = 109
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "O0011 - Lista de Organica (MALIB) por Establecimiento" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCO110.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCO110x.rpt"
            PARAMETRO = 110
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        'MOVIMIENTOS
        If NOMBRE_REPORTE = "M0001 - Resumen de Altas y Bajas por Mes" Then
            '''''''''''''''''''''''''''''''''''' AGREGAR SELECCION OBLIGATORIA, PARAMETRO APLICAR = TRUE
            If CADENA_LISTAS = "" Then
                CADENA_LISTAS = "({HISTORICO_DE_MOVIMIENTO.APLICAR} = TRUE)"
            Else
                CADENA_LISTAS = CADENA_LISTAS & " AND ({HISTORICO_DE_MOVIMIENTO.APLICAR} = TRUE)"
            End If
            SELECCION = CADENA_LISTAS
            ''''''''''''''''''''''''''''''''''''
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCM100.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCM100x.rpt"
            PARAMETRO = 111
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "M0002 - Resumen de Altas y Bajas por Niveles Organizacionales" Then
            '''''''''''''''''''''''''''''''''''' AGREGAR SELECCION OBLIGATORIA, PARAMETRO APLICAR = TRUE
            If CADENA_LISTAS = "" Then
                CADENA_LISTAS = "({HISTORICO_DE_MOVIMIENTO.APLICAR} = TRUE)"
            Else
                CADENA_LISTAS = CADENA_LISTAS & " AND ({HISTORICO_DE_MOVIMIENTO.APLICAR} = TRUE)"
            End If
            SELECCION = CADENA_LISTAS
            ''''''''''''''''''''''''''''''''''''
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCM101.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCM101x.rpt"
            PARAMETRO = 112
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "M0003 - Resumen de Altas y Bajas por Cargos" Then
            '''''''''''''''''''''''''''''''''''' AGREGAR SELECCION OBLIGATORIA, PARAMETRO APLICAR = TRUE
            If CADENA_LISTAS = "" Then
                CADENA_LISTAS = "({HISTORICO_DE_MOVIMIENTO.APLICAR} = TRUE)"
            Else
                CADENA_LISTAS = CADENA_LISTAS & " AND ({HISTORICO_DE_MOVIMIENTO.APLICAR} = TRUE)"
            End If
            SELECCION = CADENA_LISTAS
            ''''''''''''''''''''''''''''''''''''
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCM102.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCM102x.rpt"
            PARAMETRO = 113
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "M0004 - Resumen de Movimientos por Niveles Organizacionales y Rango de Fechas" Then
            '''''''''''''''''''''''''''''''''''' AGREGAR SELECCION OBLIGATORIA, PARAMETRO APLICAR = TRUE
            If CADENA_LISTAS = "" Then
                CADENA_LISTAS = "({HISTORICO_DE_MOVIMIENTO.APLICAR} = TRUE)"
            Else
                CADENA_LISTAS = CADENA_LISTAS & " AND ({HISTORICO_DE_MOVIMIENTO.APLICAR} = TRUE)"
            End If
            SELECCION = CADENA_LISTAS
            ''''''''''''''''''''''''''''''''''''
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCM103.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCM103x.rpt"
            PARAMETRO = 114
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "M0005 - Resumen de Movimientos por Cargos y Rango de Fechas" Then
            '''''''''''''''''''''''''''''''''''' AGREGAR SELECCION OBLIGATORIA, PARAMETRO APLICAR = TRUE
            If CADENA_LISTAS = "" Then
                CADENA_LISTAS = "({HISTORICO_DE_MOVIMIENTO.APLICAR} = TRUE)"
            Else
                CADENA_LISTAS = CADENA_LISTAS & " AND ({HISTORICO_DE_MOVIMIENTO.APLICAR} = TRUE)"
            End If
            SELECCION = CADENA_LISTAS
            ''''''''''''''''''''''''''''''''''''
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCM104.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCM104x.rpt"
            PARAMETRO = 115
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "M0006 - Resumen de Movimientos por Causales Reales y Cargos" Then
            '''''''''''''''''''''''''''''''''''' AGREGAR SELECCION OBLIGATORIA, PARAMETRO APLICAR = TRUE
            If CADENA_LISTAS = "" Then
                CADENA_LISTAS = "({HISTORICO_DE_MOVIMIENTO.APLICAR} = TRUE)"
            Else
                CADENA_LISTAS = CADENA_LISTAS & " AND ({HISTORICO_DE_MOVIMIENTO.APLICAR} = TRUE)"
            End If
            SELECCION = CADENA_LISTAS
            ''''''''''''''''''''''''''''''''''''
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCM105.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCM105x.rpt"
            PARAMETRO = 116
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "M0007 - Resumen de Movimientos por Causales Legales y Cargos" Then
            '''''''''''''''''''''''''''''''''''' AGREGAR SELECCION OBLIGATORIA, PARAMETRO APLICAR = TRUE
            If CADENA_LISTAS = "" Then
                CADENA_LISTAS = "({HISTORICO_DE_MOVIMIENTO.APLICAR} = TRUE)"
            Else
                CADENA_LISTAS = CADENA_LISTAS & " AND ({HISTORICO_DE_MOVIMIENTO.APLICAR} = TRUE)"
            End If
            SELECCION = CADENA_LISTAS
            ''''''''''''''''''''''''''''''''''''
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCM106.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCM106x.rpt"
            PARAMETRO = 117
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "M0008 - Resumen de Movimientos por Causales Reales y Niveles Organizacionales" Then
            '''''''''''''''''''''''''''''''''''' AGREGAR SELECCION OBLIGATORIA, PARAMETRO APLICAR = TRUE
            If CADENA_LISTAS = "" Then
                CADENA_LISTAS = "({HISTORICO_DE_MOVIMIENTO.APLICAR} = TRUE)"
            Else
                CADENA_LISTAS = CADENA_LISTAS & " AND ({HISTORICO_DE_MOVIMIENTO.APLICAR} = TRUE)"
            End If
            SELECCION = CADENA_LISTAS
            ''''''''''''''''''''''''''''''''''''
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCM107.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCM107x.rpt"
            PARAMETRO = 118
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "M0009 - Resumen de Movimientos por Causales Legales y Niveles Organizacionales" Then
            '''''''''''''''''''''''''''''''''''' AGREGAR SELECCION OBLIGATORIA, PARAMETRO APLICAR = TRUE
            If CADENA_LISTAS = "" Then
                CADENA_LISTAS = "({HISTORICO_DE_MOVIMIENTO.APLICAR} = TRUE)"
            Else
                CADENA_LISTAS = CADENA_LISTAS & " AND ({HISTORICO_DE_MOVIMIENTO.APLICAR} = TRUE)"
            End If
            SELECCION = CADENA_LISTAS
            ''''''''''''''''''''''''''''''''''''
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCM108.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCM108x.rpt"
            PARAMETRO = 119
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "M0010 - Lista de Movimientos con sus Detalles" Then
            '''''''''''''''''''''''''''''''''''' AGREGAR SELECCION OBLIGATORIA, PARAMETRO APLICAR = TRUE
            If CADENA_LISTAS = "" Then
                CADENA_LISTAS = "({HISTORICO_DE_MOVIMIENTO.APLICAR} = TRUE)"
            Else
                CADENA_LISTAS = CADENA_LISTAS & " AND ({HISTORICO_DE_MOVIMIENTO.APLICAR} = TRUE)"
            End If
            SELECCION = CADENA_LISTAS
            ''''''''''''''''''''''''''''''''''''
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCM109.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCM109x.rpt"
            PARAMETRO = 120
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "M0011 - Resumen de Movimientos por Causales Reales y Rango de Fechas" Then
            '''''''''''''''''''''''''''''''''''' AGREGAR SELECCION OBLIGATORIA, PARAMETRO APLICAR = TRUE
            If CADENA_LISTAS = "" Then
                CADENA_LISTAS = "({HISTORICO_DE_MOVIMIENTO.APLICAR} = TRUE)"
            Else
                CADENA_LISTAS = CADENA_LISTAS & " AND ({HISTORICO_DE_MOVIMIENTO.APLICAR} = TRUE)"
            End If
            SELECCION = CADENA_LISTAS
            ''''''''''''''''''''''''''''''''''''
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCM110.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCM110x.rpt"
            PARAMETRO = 151
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        'ASPIRANTE
        If NOMBRE_REPORTE = "A0001 - Lista de Aspirantes" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCA100.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCA100x.rpt"
            PARAMETRO = 121
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        'EXPEDIENTE
        If NOMBRE_REPORTE = "E0001 - Expediente de Empleado (ACTIVOS - INACTIVOS)" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCE100.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCE100x.rpt"
            PARAMETRO = 122
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "E0002 - Empleados por Antiguedad (Edad, Antiguedad PAME, Antiguedad EN)" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCE101.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCE101x.rpt"
            PARAMETRO = 123
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "E0003 - Lista de Empleados por Clasificacion" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCE102.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCE102x.rpt"
            PARAMETRO = 124
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "E0004 - Historico de Condecoraciones por Empleado" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCE103.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCE103x.rpt"
            PARAMETRO = 125
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "E0005 - Curva y Talla de Empleados por Niveles Organizacionales" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCE104.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCE104x.rpt"
            PARAMETRO = 126
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "E0006 - Lista de Empleados por Funcionalidad" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCE105.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCE105x.rpt"
            PARAMETRO = 127
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "E0007 - Contratos y Evaluaciones por Empleado" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCE106.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCE106x.rpt"
            PARAMETRO = 128
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "E0008 - Eventualidades" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCE107.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCE107x.rpt"
            PARAMETRO = 129
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "E0013 - Aplicacion del reglamento disciplinario" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCE112.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCE112x.rpt"
            PARAMETRO = 129
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "E0014 - Documentos Digitalizados por Colaborador" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCE113.rpt"
            SELECCION = "(({MAESTRO_DE_CARGOS.ID_SITUACION} = 1))"
            NOMBRE_DEL_REPORTE_EX = "SPYCE113x.rpt"
            PARAMETRO = 152
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "E0015 - Lista de Colaboradores con su Foto" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCE114.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCE114x.rpt"
            PARAMETRO = 154
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        'MEDICOS ACREDITADOS
        If NOMBRE_REPORTE = "C0001 - Lista de Medicos Acreditados" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCC100.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCC100x.rpt"
            PARAMETRO = 134
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        'BECADOS
        If NOMBRE_REPORTE = "B0001 - Lista de Becados" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCB100.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCB100x.rpt"
            PARAMETRO = 135
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        'SITUACIONES DEL PERSONAL
        If NOMBRE_REPORTE = "S0001 - Marcas por Empleado (ACTIVOS - INACTIVOS)" Then
            If Me.Lista135.Items.Count = 0 Then
                MsgBox("Debe seleccionar el Empleado del que desea visualizar las Marcas", vbInformation, "Mensaje del Sistema")
                Exit Sub
            End If
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            Call BUSCAR_EMPLEADO_ACTIVO_NO_ACTIVO()
            If ACTIVO = 1 Then
                NOMBRE_DEL_REPORTE_CR = "SPYCS100.rpt"
                NOMBRE_DEL_REPORTE_EX = "SPYCS100x.rpt"
                PARAMETRO = 136
            Else
                NOMBRE_DEL_REPORTE_CR = "SPYCS101.rpt"
                NOMBRE_DEL_REPORTE_EX = "SPYCS101x.rpt"
                PARAMETRO = 137
            End If
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "S0002 - Marcas por Niveles Organizacionales" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCS102.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCS102x.rpt"
            PARAMETRO = 138
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        'If NOMBRE_REPORTE = "S0003 - Formato de Solicitud de Vacaciones" Then
        '    SELECCION_PARAMETRO = "*"
        '    USUARIO_IMPRIME = isuCUENTA
        '    NOMBRE_DEL_REPORTE_CR = "SPYCS103.rpt"
        '    NOMBRE_DEL_REPORTE_EX = "SPYCS103x.rpt"
        '    PARAMETRO = 139
        '    Call PROCESO_IMRPIMIR()
        '    GoTo SALTO
        'End If
        If NOMBRE_REPORTE = "S0004 - Cumpleañeros" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCS104.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCS104x.rpt"
            PARAMETRO = 140
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "S0005 - Subsidios por Niveles Organizacionales" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCS105.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCS105x.rpt"
            PARAMETRO = 141
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "S0006 - Detalles de Subsidios por Niveles Organizacionales" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCS106.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCS106x.rpt"
            PARAMETRO = 142
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "S0007 - Situaciones del Personal por Rango de Fechas" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCS107.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCS107x.rpt"
            PARAMETRO = 143
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "S0008 - Estado Vacacional por Rango de Fechas (Detalles)" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            SELECCION = SELECCION & " AND ({MAESTRO_DE_PERSONAS.CALCULO_VACACIONAL} = TRUE)"
            NOMBRE_DEL_REPORTE_CR = "SPYCS108.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCS108x.rpt"
            PARAMETRO = 144
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "S0009 - Estado Vacacional por Rango de Fechas (Niveles Organizacionales)" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            SELECCION = SELECCION & " AND ({MAESTRO_DE_PERSONAS.CALCULO_VACACIONAL} = TRUE)"
            NOMBRE_DEL_REPORTE_CR = "SPYCS109.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCS109x.rpt"
            PARAMETRO = 145
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "S0010 - Indicadores de < Cantidad de Empleados por Tipo de Situacion A >" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCS110.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCS110x.rpt"
            PARAMETRO = 146
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "S0011 - Indicadores de < Cantidad de Empleados por Tipo de Situacion B >" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCS111.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCS111x.rpt"
            PARAMETRO = 147
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "S0012 - Indicadores de < Cantidad de Dias por Empleado y Tipo de Situacion >" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCS112.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCS112x.rpt"
            PARAMETRO = 148
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        If NOMBRE_REPORTE = "S0013 - Indicadores de < Lista de Empleados por Tipo de Situacion >" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCS113.rpt"
            NOMBRE_DEL_REPORTE_EX = "SPYCS113x.rpt"
            PARAMETRO = 149
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        'HIGIENE Y SEGURIDAD
        If NOMBRE_REPORTE = "D0002 - Lista de Colaboradores con Inmunizaciones" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            NOMBRE_DEL_REPORTE_CR = "SPYCD002.rpt"
            SELECCION = "(({MAESTRO_DE_CARGOS.ID_SITUACION} = 1))"
            NOMBRE_DEL_REPORTE_EX = "SPYCD002x.rpt"
            PARAMETRO = 153
            Call PROCESO_IMRPIMIR()
            GoTo SALTO
        End If
SALTO:
    End Sub
    Dim ACTIVO As Integer
    Private Sub BUSCAR_EMPLEADO_ACTIVO_NO_ACTIVO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [MAESTRO DE PERSONAS] WHERE CODIGO = '" & vCOD & "'", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                ACTIVO = MiDataTable.Rows(0).Item(46).ToString
                If ACTIVO = 1 Then  'ACTIVO
                    VALOR01 = MiDataTable.Rows(0).Item(4).ToString
                    VALOR02 = MiDataTable.Rows(0).Item(5).ToString & " " & MiDataTable.Rows(0).Item(6).ToString
                    VALOR04 = Mid(Me.DateTimePicker1.Value, 1, 10)
                    VALOR05 = Mid(Me.DateTimePicker2.Value, 1, 10)
                    Call BUSCAR_CARGO_ACTUAL
                End If

                If ACTIVO = 2 Then  'NO ACTIVO
                    VALOR01 = MiDataTable.Rows(0).Item(4).ToString
                    VALOR02 = MiDataTable.Rows(0).Item(5).ToString & " " & MiDataTable.Rows(0).Item(6).ToString
                    VALOR03 = "."
                    VALOR04 = Mid(Me.DateTimePicker1.Value, 1, 10)
                    VALOR05 = Mid(Me.DateTimePicker2.Value, 1, 10)
                    VALOR06 = "."
                End If
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub BUSCAR_CARGO_ACTUAL()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [VISTA MAESTRO DE CARGOS] WHERE CODIGO = '" & vCOD & "'", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                VALOR03 = MiDataTable.Rows(0).Item(32).ToString
                Dim cadenaUBIC As String
                cadenaUBIC = MiDataTable.Rows(0).Item(4).ToString
                If MiDataTable.Rows(0).Item(7).ToString = "" Then
                    GoTo SALIR
                Else
                    cadenaUBIC = cadenaUBIC & "\ " & MiDataTable.Rows(0).Item(7).ToString
                End If
                If MiDataTable.Rows(0).Item(10).ToString = "" Then
                    GoTo SALIR
                Else
                    cadenaUBIC = cadenaUBIC & "\ " & MiDataTable.Rows(0).Item(10).ToString
                End If
                If MiDataTable.Rows(0).Item(13).ToString = "" Then
                    GoTo SALIR
                Else
                    cadenaUBIC = cadenaUBIC & "\ " & MiDataTable.Rows(0).Item(13).ToString
                End If
                If MiDataTable.Rows(0).Item(16).ToString = "" Then
                    GoTo SALIR
                Else
                    cadenaUBIC = cadenaUBIC & "\ " & MiDataTable.Rows(0).Item(16).ToString
                End If
                If MiDataTable.Rows(0).Item(19).ToString = "" Then
                    GoTo SALIR
                Else
                    cadenaUBIC = cadenaUBIC & "\ " & MiDataTable.Rows(0).Item(19).ToString
                End If
                If MiDataTable.Rows(0).Item(22).ToString = "" Then
                    GoTo SALIR
                Else
                    cadenaUBIC = cadenaUBIC & "\ " & MiDataTable.Rows(0).Item(22).ToString
                End If
                If MiDataTable.Rows(0).Item(25).ToString = "" Then
                    GoTo SALIR
                Else
                    cadenaUBIC = cadenaUBIC & "\ " & MiDataTable.Rows(0).Item(25).ToString
                End If
SALIR:
                VALOR06 = cadenaUBIC                                      'UB

            Else
                VALOR06 = "."
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim ARMAR_CADENA_LISTA_1 As String
    Dim ARMAR_CADENA_LISTA_2 As String
    Dim ARMAR_CADENA_LISTA_3 As String
    Dim ARMAR_CADENA_LISTA_4 As String
    Dim ARMAR_CADENA_LISTA_5 As String
    Dim ARMAR_CADENA_LISTA_6 As String
    Dim ARMAR_CADENA_LISTA_7 As String
    Dim ARMAR_CADENA_LISTA_8 As String
    Dim ARMAR_CADENA_LISTA_9 As String
    Dim ARMAR_CADENA_LISTA_10 As String
    Dim ARMAR_CADENA_LISTA_11 As String
    Dim ARMAR_CADENA_LISTA_12 As String
    Dim ARMAR_CADENA_LISTA_13 As String
    Dim ARMAR_CADENA_LISTA_14 As String
    Dim ARMAR_CADENA_LISTA_15 As String
    Dim ARMAR_CADENA_LISTA_16 As String
    Dim ARMAR_CADENA_LISTA_17 As String
    Dim ARMAR_CADENA_LISTA_18 As String
    Dim ARMAR_CADENA_LISTA_19 As String
    Dim ARMAR_CADENA_LISTA_20 As String
    Dim ARMAR_CADENA_LISTA_21 As String
    Dim ARMAR_CADENA_LISTA_22 As String
    Dim ARMAR_CADENA_LISTA_23 As String
    Dim ARMAR_CADENA_LISTA_24 As String
    Dim ARMAR_CADENA_LISTA_25 As String
    Dim ARMAR_CADENA_LISTA_26 As String
    Dim ARMAR_CADENA_LISTA_27 As String
    Dim ARMAR_CADENA_LISTA_28 As String
    Dim ARMAR_CADENA_LISTA_29 As String
    Dim ARMAR_CADENA_LISTA_30 As String
    Dim ARMAR_CADENA_LISTA_31 As String
    Dim ARMAR_CADENA_LISTA_32 As String
    Dim ARMAR_CADENA_LISTA_33 As String
    Dim ARMAR_CADENA_LISTA_34 As String
    Dim ARMAR_CADENA_LISTA_35 As String
    Dim ARMAR_CADENA_LISTA_36 As String
    Dim ARMAR_CADENA_LISTA_37 As String
    Dim ARMAR_CADENA_LISTA_38 As String
    Dim ARMAR_CADENA_LISTA_39 As String
    Dim ARMAR_CADENA_LISTA_40 As String
    Dim ARMAR_CADENA_LISTA_41 As String
    Dim ARMAR_CADENA_LISTA_42 As String
    Dim ARMAR_CADENA_LISTA_43 As String
    Dim CADENA_LISTAS As String
    Private Sub ARMAR_CADENA_GENERAL()
        ARMAR_CADENA_LISTA_1 = "" : ARMAR_CADENA_LISTA_2 = "" : ARMAR_CADENA_LISTA_3 = "" : ARMAR_CADENA_LISTA_4 = "" : ARMAR_CADENA_LISTA_5 = ""
        ARMAR_CADENA_LISTA_6 = "" : ARMAR_CADENA_LISTA_7 = "" : ARMAR_CADENA_LISTA_8 = "" : ARMAR_CADENA_LISTA_9 = "" : ARMAR_CADENA_LISTA_10 = ""
        ARMAR_CADENA_LISTA_11 = "" : ARMAR_CADENA_LISTA_12 = "" : ARMAR_CADENA_LISTA_13 = "" : ARMAR_CADENA_LISTA_14 = "" : ARMAR_CADENA_LISTA_15 = ""
        ARMAR_CADENA_LISTA_16 = "" : ARMAR_CADENA_LISTA_17 = "" : ARMAR_CADENA_LISTA_18 = "" : ARMAR_CADENA_LISTA_19 = "" : ARMAR_CADENA_LISTA_20 = ""
        ARMAR_CADENA_LISTA_21 = "" : ARMAR_CADENA_LISTA_22 = "" : ARMAR_CADENA_LISTA_23 = "" : ARMAR_CADENA_LISTA_24 = "" : ARMAR_CADENA_LISTA_25 = ""
        ARMAR_CADENA_LISTA_26 = "" : ARMAR_CADENA_LISTA_27 = "" : ARMAR_CADENA_LISTA_28 = "" : ARMAR_CADENA_LISTA_29 = "" : ARMAR_CADENA_LISTA_30 = ""
        ARMAR_CADENA_LISTA_31 = "" : ARMAR_CADENA_LISTA_32 = "" : ARMAR_CADENA_LISTA_33 = "" : ARMAR_CADENA_LISTA_34 = "" : ARMAR_CADENA_LISTA_35 = ""
        ARMAR_CADENA_LISTA_36 = "" : ARMAR_CADENA_LISTA_37 = "" : ARMAR_CADENA_LISTA_38 = "" : ARMAR_CADENA_LISTA_39 = "" : ARMAR_CADENA_LISTA_40 = ""
        ARMAR_CADENA_LISTA_41 = "" : ARMAR_CADENA_LISTA_42 = "" : ARMAR_CADENA_LISTA_43 = ""
        CADENA_LISTAS = ""
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Me.Lista101.Items.Count <> 0 Then : For I = 0 To Me.Lista101.Items.Count - 1 : If ARMAR_CADENA_LISTA_1 = "" Then : ARMAR_CADENA_LISTA_1 = Me.Lista101.Items(I).ToString : Else : ARMAR_CADENA_LISTA_1 = ARMAR_CADENA_LISTA_1 & " " & Me.Lista101.Items(I).ToString : End If : Next : End If
        If Me.Lista102.Items.Count <> 0 Then : For I = 0 To Me.Lista102.Items.Count - 1 : If ARMAR_CADENA_LISTA_2 = "" Then : ARMAR_CADENA_LISTA_2 = Me.Lista102.Items(I).ToString : Else : ARMAR_CADENA_LISTA_2 = ARMAR_CADENA_LISTA_2 & " " & Me.Lista102.Items(I).ToString : End If : Next : End If
        If Me.Lista103.Items.Count <> 0 Then : For I = 0 To Me.Lista103.Items.Count - 1 : If ARMAR_CADENA_LISTA_3 = "" Then : ARMAR_CADENA_LISTA_3 = Me.Lista103.Items(I).ToString : Else : ARMAR_CADENA_LISTA_3 = ARMAR_CADENA_LISTA_3 & " " & Me.Lista103.Items(I).ToString : End If : Next : End If
        If Me.Lista104.Items.Count <> 0 Then : For I = 0 To Me.Lista104.Items.Count - 1 : If ARMAR_CADENA_LISTA_4 = "" Then : ARMAR_CADENA_LISTA_4 = Me.Lista104.Items(I).ToString : Else : ARMAR_CADENA_LISTA_4 = ARMAR_CADENA_LISTA_4 & " " & Me.Lista104.Items(I).ToString : End If : Next : End If
        If Me.Lista105.Items.Count <> 0 Then : For I = 0 To Me.Lista105.Items.Count - 1 : If ARMAR_CADENA_LISTA_5 = "" Then : ARMAR_CADENA_LISTA_5 = Me.Lista105.Items(I).ToString : Else : ARMAR_CADENA_LISTA_5 = ARMAR_CADENA_LISTA_5 & " " & Me.Lista105.Items(I).ToString : End If : Next : End If
        If Me.Lista106.Items.Count <> 0 Then : For I = 0 To Me.Lista106.Items.Count - 1 : If ARMAR_CADENA_LISTA_6 = "" Then : ARMAR_CADENA_LISTA_6 = Me.Lista106.Items(I).ToString : Else : ARMAR_CADENA_LISTA_6 = ARMAR_CADENA_LISTA_6 & " " & Me.Lista106.Items(I).ToString : End If : Next : End If
        If Me.Lista107.Items.Count <> 0 Then : For I = 0 To Me.Lista107.Items.Count - 1 : If ARMAR_CADENA_LISTA_7 = "" Then : ARMAR_CADENA_LISTA_7 = Me.Lista107.Items(I).ToString : Else : ARMAR_CADENA_LISTA_7 = ARMAR_CADENA_LISTA_7 & " " & Me.Lista107.Items(I).ToString : End If : Next : End If
        If Me.Lista108.Items.Count <> 0 Then : For I = 0 To Me.Lista108.Items.Count - 1 : If ARMAR_CADENA_LISTA_8 = "" Then : ARMAR_CADENA_LISTA_8 = Me.Lista108.Items(I).ToString : Else : ARMAR_CADENA_LISTA_8 = ARMAR_CADENA_LISTA_8 & " " & Me.Lista108.Items(I).ToString : End If : Next : End If
        If Me.Lista109.Items.Count <> 0 Then : For I = 0 To Me.Lista109.Items.Count - 1 : If ARMAR_CADENA_LISTA_9 = "" Then : ARMAR_CADENA_LISTA_9 = Me.Lista109.Items(I).ToString : Else : ARMAR_CADENA_LISTA_9 = ARMAR_CADENA_LISTA_9 & " " & Me.Lista109.Items(I).ToString : End If : Next : End If
        If Me.Lista110.Items.Count <> 0 Then : For I = 0 To Me.Lista110.Items.Count - 1 : If ARMAR_CADENA_LISTA_10 = "" Then : ARMAR_CADENA_LISTA_10 = Me.Lista110.Items(I).ToString : Else : ARMAR_CADENA_LISTA_10 = ARMAR_CADENA_LISTA_10 & " " & Me.Lista110.Items(I).ToString : End If : Next : End If
        If Me.Lista111.Items.Count <> 0 Then : For I = 0 To Me.Lista111.Items.Count - 1 : If ARMAR_CADENA_LISTA_11 = "" Then : ARMAR_CADENA_LISTA_11 = Me.Lista111.Items(I).ToString : Else : ARMAR_CADENA_LISTA_11 = ARMAR_CADENA_LISTA_11 & " " & Me.Lista111.Items(I).ToString : End If : Next : End If
        If Me.Lista112.Items.Count <> 0 Then : For I = 0 To Me.Lista112.Items.Count - 1 : If ARMAR_CADENA_LISTA_12 = "" Then : ARMAR_CADENA_LISTA_12 = Me.Lista112.Items(I).ToString : Else : ARMAR_CADENA_LISTA_12 = ARMAR_CADENA_LISTA_12 & " " & Me.Lista112.Items(I).ToString : End If : Next : End If
        If Me.Lista113.Items.Count <> 0 Then : For I = 0 To Me.Lista113.Items.Count - 1 : If ARMAR_CADENA_LISTA_13 = "" Then : ARMAR_CADENA_LISTA_13 = Me.Lista113.Items(I).ToString : Else : ARMAR_CADENA_LISTA_13 = ARMAR_CADENA_LISTA_13 & " " & Me.Lista113.Items(I).ToString : End If : Next : End If
        If Me.Lista114.Items.Count <> 0 Then : For I = 0 To Me.Lista114.Items.Count - 1 : If ARMAR_CADENA_LISTA_14 = "" Then : ARMAR_CADENA_LISTA_14 = Me.Lista114.Items(I).ToString : Else : ARMAR_CADENA_LISTA_14 = ARMAR_CADENA_LISTA_14 & " " & Me.Lista114.Items(I).ToString : End If : Next : End If
        If Me.Lista115.Items.Count <> 0 Then : For I = 0 To Me.Lista115.Items.Count - 1 : If ARMAR_CADENA_LISTA_15 = "" Then : ARMAR_CADENA_LISTA_15 = Me.Lista115.Items(I).ToString : Else : ARMAR_CADENA_LISTA_15 = ARMAR_CADENA_LISTA_15 & " " & Me.Lista115.Items(I).ToString : End If : Next : End If
        If Me.Lista116.Items.Count <> 0 Then : For I = 0 To Me.Lista116.Items.Count - 1 : If ARMAR_CADENA_LISTA_16 = "" Then : ARMAR_CADENA_LISTA_16 = Me.Lista116.Items(I).ToString : Else : ARMAR_CADENA_LISTA_16 = ARMAR_CADENA_LISTA_16 & " " & Me.Lista116.Items(I).ToString : End If : Next : End If
        If Me.Lista117.Items.Count <> 0 Then : For I = 0 To Me.Lista117.Items.Count - 1 : If ARMAR_CADENA_LISTA_17 = "" Then : ARMAR_CADENA_LISTA_17 = Me.Lista117.Items(I).ToString : Else : ARMAR_CADENA_LISTA_17 = ARMAR_CADENA_LISTA_17 & " " & Me.Lista117.Items(I).ToString : End If : Next : End If
        If Me.Lista118.Items.Count <> 0 Then : For I = 0 To Me.Lista118.Items.Count - 1 : If ARMAR_CADENA_LISTA_18 = "" Then : ARMAR_CADENA_LISTA_18 = Me.Lista118.Items(I).ToString : Else : ARMAR_CADENA_LISTA_18 = ARMAR_CADENA_LISTA_18 & " " & Me.Lista118.Items(I).ToString : End If : Next : End If
        If Me.Lista119.Items.Count <> 0 Then : For I = 0 To Me.Lista119.Items.Count - 1 : If ARMAR_CADENA_LISTA_19 = "" Then : ARMAR_CADENA_LISTA_19 = Me.Lista119.Items(I).ToString : Else : ARMAR_CADENA_LISTA_19 = ARMAR_CADENA_LISTA_19 & " " & Me.Lista119.Items(I).ToString : End If : Next : End If
        If Me.Lista120.Items.Count <> 0 Then : For I = 0 To Me.Lista120.Items.Count - 1 : If ARMAR_CADENA_LISTA_20 = "" Then : ARMAR_CADENA_LISTA_20 = Me.Lista120.Items(I).ToString : Else : ARMAR_CADENA_LISTA_20 = ARMAR_CADENA_LISTA_20 & " " & Me.Lista120.Items(I).ToString : End If : Next : End If
        If Me.Lista121.Items.Count <> 0 Then : For I = 0 To Me.Lista121.Items.Count - 1 : If ARMAR_CADENA_LISTA_21 = "" Then : ARMAR_CADENA_LISTA_21 = Me.Lista121.Items(I).ToString : Else : ARMAR_CADENA_LISTA_21 = ARMAR_CADENA_LISTA_21 & " " & Me.Lista121.Items(I).ToString : End If : Next : End If
        If Me.Lista122.Items.Count <> 0 Then : For I = 0 To Me.Lista122.Items.Count - 1 : If ARMAR_CADENA_LISTA_22 = "" Then : ARMAR_CADENA_LISTA_22 = Me.Lista122.Items(I).ToString : Else : ARMAR_CADENA_LISTA_22 = ARMAR_CADENA_LISTA_22 & " " & Me.Lista122.Items(I).ToString : End If : Next : End If
        If Me.Lista123.Items.Count <> 0 Then : For I = 0 To Me.Lista123.Items.Count - 1 : If ARMAR_CADENA_LISTA_23 = "" Then : ARMAR_CADENA_LISTA_23 = Me.Lista123.Items(I).ToString : Else : ARMAR_CADENA_LISTA_23 = ARMAR_CADENA_LISTA_23 & " " & Me.Lista123.Items(I).ToString : End If : Next : End If
        If Me.Lista124.Items.Count <> 0 Then : For I = 0 To Me.Lista124.Items.Count - 1 : If ARMAR_CADENA_LISTA_24 = "" Then : ARMAR_CADENA_LISTA_24 = Me.Lista124.Items(I).ToString : Else : ARMAR_CADENA_LISTA_24 = ARMAR_CADENA_LISTA_24 & " " & Me.Lista124.Items(I).ToString : End If : Next : End If
        If Me.Lista125.Items.Count <> 0 Then : For I = 0 To Me.Lista125.Items.Count - 1 : If ARMAR_CADENA_LISTA_25 = "" Then : ARMAR_CADENA_LISTA_25 = Me.Lista125.Items(I).ToString : Else : ARMAR_CADENA_LISTA_25 = ARMAR_CADENA_LISTA_25 & " " & Me.Lista125.Items(I).ToString : End If : Next : End If
        If Me.Lista126.Items.Count <> 0 Then : For I = 0 To Me.Lista126.Items.Count - 1 : If ARMAR_CADENA_LISTA_26 = "" Then : ARMAR_CADENA_LISTA_26 = Me.Lista126.Items(I).ToString : Else : ARMAR_CADENA_LISTA_26 = ARMAR_CADENA_LISTA_26 & " " & Me.Lista126.Items(I).ToString : End If : Next : End If
        If Me.Lista127.Items.Count <> 0 Then : For I = 0 To Me.Lista127.Items.Count - 1 : If ARMAR_CADENA_LISTA_27 = "" Then : ARMAR_CADENA_LISTA_27 = Me.Lista127.Items(I).ToString : Else : ARMAR_CADENA_LISTA_27 = ARMAR_CADENA_LISTA_27 & " " & Me.Lista127.Items(I).ToString : End If : Next : End If
        If Me.Lista128.Items.Count <> 0 Then : For I = 0 To Me.Lista128.Items.Count - 1 : If ARMAR_CADENA_LISTA_28 = "" Then : ARMAR_CADENA_LISTA_28 = Me.Lista128.Items(I).ToString : Else : ARMAR_CADENA_LISTA_28 = ARMAR_CADENA_LISTA_28 & " " & Me.Lista128.Items(I).ToString : End If : Next : End If
        If Me.Lista129.Items.Count <> 0 Then : For I = 0 To Me.Lista129.Items.Count - 1 : If ARMAR_CADENA_LISTA_29 = "" Then : ARMAR_CADENA_LISTA_29 = Me.Lista129.Items(I).ToString : Else : ARMAR_CADENA_LISTA_29 = ARMAR_CADENA_LISTA_29 & " " & Me.Lista129.Items(I).ToString : End If : Next : End If
        If Me.Lista130.Items.Count <> 0 Then : For I = 0 To Me.Lista130.Items.Count - 1 : If ARMAR_CADENA_LISTA_30 = "" Then : ARMAR_CADENA_LISTA_30 = Me.Lista130.Items(I).ToString : Else : ARMAR_CADENA_LISTA_30 = ARMAR_CADENA_LISTA_30 & " " & Me.Lista130.Items(I).ToString : End If : Next : End If
        If Me.Lista131.Items.Count <> 0 Then : For I = 0 To Me.Lista131.Items.Count - 1 : If ARMAR_CADENA_LISTA_31 = "" Then : ARMAR_CADENA_LISTA_31 = Me.Lista131.Items(I).ToString : Else : ARMAR_CADENA_LISTA_31 = ARMAR_CADENA_LISTA_31 & " " & Me.Lista131.Items(I).ToString : End If : Next : End If
        If Me.Lista132.Items.Count <> 0 Then : For I = 0 To Me.Lista132.Items.Count - 1 : If ARMAR_CADENA_LISTA_32 = "" Then : ARMAR_CADENA_LISTA_32 = Me.Lista132.Items(I).ToString : Else : ARMAR_CADENA_LISTA_32 = ARMAR_CADENA_LISTA_32 & " " & Me.Lista132.Items(I).ToString : End If : Next : End If
        If Me.Lista133.Items.Count <> 0 Then : For I = 0 To Me.Lista133.Items.Count - 1 : If ARMAR_CADENA_LISTA_33 = "" Then : ARMAR_CADENA_LISTA_33 = Me.Lista133.Items(I).ToString : Else : ARMAR_CADENA_LISTA_33 = ARMAR_CADENA_LISTA_33 & " " & Me.Lista133.Items(I).ToString : End If : Next : End If
        If Me.Lista134.Items.Count <> 0 Then : For I = 0 To Me.Lista134.Items.Count - 1 : If ARMAR_CADENA_LISTA_34 = "" Then : ARMAR_CADENA_LISTA_34 = Me.Lista134.Items(I).ToString : Else : ARMAR_CADENA_LISTA_34 = ARMAR_CADENA_LISTA_34 & " " & Me.Lista134.Items(I).ToString : End If : Next : End If
        If Me.Lista135.Items.Count <> 0 Then : For I = 0 To Me.Lista135.Items.Count - 1 : If ARMAR_CADENA_LISTA_35 = "" Then : ARMAR_CADENA_LISTA_35 = Me.Lista135.Items(I).ToString : Else : ARMAR_CADENA_LISTA_35 = ARMAR_CADENA_LISTA_35 & " " & Me.Lista135.Items(I).ToString : End If : Next : End If
        If Me.Lista136.Items.Count <> 0 Then : For I = 0 To Me.Lista136.Items.Count - 1 : If ARMAR_CADENA_LISTA_36 = "" Then : ARMAR_CADENA_LISTA_36 = Me.Lista136.Items(I).ToString : Else : ARMAR_CADENA_LISTA_36 = ARMAR_CADENA_LISTA_36 & " " & Me.Lista136.Items(I).ToString : End If : Next : End If
        If Me.Lista137.Items.Count <> 0 Then : For I = 0 To Me.Lista137.Items.Count - 1 : If ARMAR_CADENA_LISTA_37 = "" Then : ARMAR_CADENA_LISTA_37 = Me.Lista137.Items(I).ToString : Else : ARMAR_CADENA_LISTA_37 = ARMAR_CADENA_LISTA_37 & " " & Me.Lista137.Items(I).ToString : End If : Next : End If
        If Me.Lista138.Items.Count <> 0 Then : For I = 0 To Me.Lista138.Items.Count - 1 : If ARMAR_CADENA_LISTA_38 = "" Then : ARMAR_CADENA_LISTA_38 = Me.Lista138.Items(I).ToString : Else : ARMAR_CADENA_LISTA_38 = ARMAR_CADENA_LISTA_38 & " " & Me.Lista138.Items(I).ToString : End If : Next : End If
        If Me.Lista139.Items.Count <> 0 Then : For I = 0 To Me.Lista139.Items.Count - 1 : If ARMAR_CADENA_LISTA_39 = "" Then : ARMAR_CADENA_LISTA_39 = Me.Lista139.Items(I).ToString : Else : ARMAR_CADENA_LISTA_39 = ARMAR_CADENA_LISTA_39 & " " & Me.Lista139.Items(I).ToString : End If : Next : End If
        If Me.Lista140.Items.Count <> 0 Then : For I = 0 To Me.Lista140.Items.Count - 1 : If ARMAR_CADENA_LISTA_40 = "" Then : ARMAR_CADENA_LISTA_40 = Me.Lista140.Items(I).ToString : Else : ARMAR_CADENA_LISTA_40 = ARMAR_CADENA_LISTA_40 & " " & Me.Lista140.Items(I).ToString : End If : Next : End If
        If Me.Lista141.Items.Count <> 0 Then : For I = 0 To Me.Lista141.Items.Count - 1 : If ARMAR_CADENA_LISTA_41 = "" Then : ARMAR_CADENA_LISTA_41 = Me.Lista141.Items(I).ToString : Else : ARMAR_CADENA_LISTA_41 = ARMAR_CADENA_LISTA_41 & " " & Me.Lista141.Items(I).ToString : End If : Next : End If
        If Me.Lista142.Items.Count <> 0 Then : For I = 0 To Me.Lista142.Items.Count - 1 : If ARMAR_CADENA_LISTA_42 = "" Then : ARMAR_CADENA_LISTA_42 = Me.Lista142.Items(I).ToString : Else : ARMAR_CADENA_LISTA_42 = ARMAR_CADENA_LISTA_42 & " " & Me.Lista142.Items(I).ToString : End If : Next : End If
        If Me.Lista143.Items.Count <> 0 Then : For I = 0 To Me.Lista143.Items.Count - 1 : If ARMAR_CADENA_LISTA_43 = "" Then : ARMAR_CADENA_LISTA_43 = Me.Lista143.Items(I).ToString : Else : ARMAR_CADENA_LISTA_43 = ARMAR_CADENA_LISTA_43 & " " & Me.Lista143.Items(I).ToString : End If : Next : End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_1 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_1 & ")" : End If : Else : If ARMAR_CADENA_LISTA_1 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_1 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_2 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_2 & ")" : End If : Else : If ARMAR_CADENA_LISTA_2 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_2 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_3 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_3 & ")" : End If : Else : If ARMAR_CADENA_LISTA_3 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_3 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_4 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_4 & ")" : End If : Else : If ARMAR_CADENA_LISTA_4 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_4 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_5 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_5 & ")" : End If : Else : If ARMAR_CADENA_LISTA_5 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_5 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_6 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_6 & ")" : End If : Else : If ARMAR_CADENA_LISTA_6 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_6 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_7 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_7 & ")" : End If : Else : If ARMAR_CADENA_LISTA_7 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_7 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_8 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_8 & ")" : End If : Else : If ARMAR_CADENA_LISTA_8 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_8 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_9 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_9 & ")" : End If : Else : If ARMAR_CADENA_LISTA_9 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_9 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_10 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_10 & ")" : End If : Else : If ARMAR_CADENA_LISTA_10 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_10 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_11 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_11 & ")" : End If : Else : If ARMAR_CADENA_LISTA_11 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_11 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_12 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_12 & ")" : End If : Else : If ARMAR_CADENA_LISTA_12 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_12 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_13 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_13 & ")" : End If : Else : If ARMAR_CADENA_LISTA_13 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_13 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_14 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_14 & ")" : End If : Else : If ARMAR_CADENA_LISTA_14 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_14 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_15 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_15 & ")" : End If : Else : If ARMAR_CADENA_LISTA_15 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_15 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_16 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_16 & ")" : End If : Else : If ARMAR_CADENA_LISTA_16 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_16 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_17 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_17 & ")" : End If : Else : If ARMAR_CADENA_LISTA_17 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_17 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_18 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_18 & ")" : End If : Else : If ARMAR_CADENA_LISTA_18 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_18 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_19 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_19 & ")" : End If : Else : If ARMAR_CADENA_LISTA_19 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_19 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_20 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_20 & ")" : End If : Else : If ARMAR_CADENA_LISTA_20 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_20 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_21 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_21 & ")" : End If : Else : If ARMAR_CADENA_LISTA_21 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_21 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_22 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_22 & ")" : End If : Else : If ARMAR_CADENA_LISTA_22 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_22 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_23 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_23 & ")" : End If : Else : If ARMAR_CADENA_LISTA_23 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_23 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_24 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_24 & ")" : End If : Else : If ARMAR_CADENA_LISTA_24 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_24 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_25 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_25 & ")" : End If : Else : If ARMAR_CADENA_LISTA_25 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_25 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_26 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_26 & ")" : End If : Else : If ARMAR_CADENA_LISTA_26 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_26 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_27 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_27 & ")" : End If : Else : If ARMAR_CADENA_LISTA_27 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_27 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_28 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_28 & ")" : End If : Else : If ARMAR_CADENA_LISTA_28 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_28 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_29 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_29 & ")" : End If : Else : If ARMAR_CADENA_LISTA_29 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_29 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_30 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_30 & ")" : End If : Else : If ARMAR_CADENA_LISTA_30 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_30 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_31 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_31 & ")" : End If : Else : If ARMAR_CADENA_LISTA_31 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_31 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_32 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_32 & ")" : End If : Else : If ARMAR_CADENA_LISTA_32 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_32 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_33 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_33 & ")" : End If : Else : If ARMAR_CADENA_LISTA_33 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_33 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_34 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_34 & ")" : End If : Else : If ARMAR_CADENA_LISTA_34 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_34 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_35 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_35 & ")" : End If : Else : If ARMAR_CADENA_LISTA_35 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_35 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_36 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_36 & ")" : End If : Else : If ARMAR_CADENA_LISTA_36 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_36 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_37 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_37 & ")" : End If : Else : If ARMAR_CADENA_LISTA_37 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_37 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_38 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_38 & ")" : End If : Else : If ARMAR_CADENA_LISTA_38 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_38 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_39 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_39 & ")" : End If : Else : If ARMAR_CADENA_LISTA_39 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_39 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_40 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_40 & ")" : End If : Else : If ARMAR_CADENA_LISTA_40 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_40 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_41 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_41 & ")" : End If : Else : If ARMAR_CADENA_LISTA_41 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_41 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_42 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_42 & ")" : End If : Else : If ARMAR_CADENA_LISTA_42 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_42 & ")" : End If : End If
        If CADENA_LISTAS = "" Then : If ARMAR_CADENA_LISTA_43 <> "" Then : CADENA_LISTAS = "(" & ARMAR_CADENA_LISTA_43 & ")" : End If : Else : If ARMAR_CADENA_LISTA_43 <> "" Then : CADENA_LISTAS = CADENA_LISTAS & " AND (" & ARMAR_CADENA_LISTA_43 & ")" : End If : End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End Sub
    Private Sub ARMAR_CADENA_FILTRO_SELECCION_REPORTE()
        FILTRO_SELECCION_EN_REPORTES = ""
        If Me.Lista_Principal.Items.Count <> 0 Then
            If Me.Lista_Principal.Items.Count <> 0 Then
                For I = 0 To Me.Lista_Principal.Items.Count - 1
                    If FILTRO_SELECCION_EN_REPORTES = "" Then
                        FILTRO_SELECCION_EN_REPORTES = Me.Lista_Principal.Items(I).ToString
                    Else
                        FILTRO_SELECCION_EN_REPORTES = FILTRO_SELECCION_EN_REPORTES & "; " & Me.Lista_Principal.Items(I).ToString
                    End If
                Next
            End If
        End If
        If FILTRO_SELECCION_EN_REPORTES = "" Then
            FILTRO_SELECCION_EN_REPORTES = "- SIN FILTRO -"
        End If
    End Sub
    Private Sub PROCESO_IMRPIMIR()
        If Me.ComboBox100.Text = "CRYSTAL" Then
            Frm_Visor_Reportes.CrystalR.ShowExportButton = True
            Frm_Visor_Reportes.CrystalR.ShowPrintButton = True
            Frm_Visor_Reportes.ShowDialog()
            GoTo SALTO
        End If
        '////////////////////////////////////////////////////////////////////////////////////////////////////
        '////////////////////////////////////////////////////////////////////////////////////////////////////
        If Me.ComboBox100.Text = "OFFICE (EXCEL\WORD)" Then
            Dim rd = New ReportDocument()
            Dim PathRPT As String = "\\Hmil-spc-apprd\e$\v3 SPYC\Excel\"
            rd.Load(PathRPT & NOMBRE_DEL_REPORTE_EX)
            rd.SetDatabaseLogon("sa", "P@$$W0RD")
            Me.Cursor = Cursors.WaitCursor
            Dim NOMBRE_ARMADO As String

            If NOMBRE_DEL_REPORTE_CR = "SPYCS103.rpt" Then
                NOMBRE_ARMADO = "tmpQUERY_" & PARAMETRO & "_" & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00") & ".doc"
                Dim objWORDOptions As PdfRtfWordFormatOptions = New PdfRtfWordFormatOptions
                'objWORDOptions.ExcelUseConstantColumnWidth = False
                rd.ExportOptions.FormatOptions = objWORDOptions
                Dim strExportFile_WORD As String = "C:\WINDOWS\TEMP\" & NOMBRE_ARMADO
                rd.ExportOptions.ExportDestinationType = ExportDestinationType.DiskFile
                rd.ExportOptions.ExportFormatType = ExportFormatType.WordForWindows
                'objWORDOptions.ExcelUseConstantColumnWidth = False
                rd.ExportOptions.FormatOptions = objWORDOptions
                Dim objOptions_WORD As DiskFileDestinationOptions = New DiskFileDestinationOptions
                objOptions_WORD.DiskFileName = strExportFile_WORD
                rd.ExportOptions.DestinationOptions = objOptions_WORD
                rd.Export()
                objOptions_WORD = Nothing
                rd = Nothing
                Process.Start("WINWORD.exe", "C:\WINDOWS\TEMP\" & NOMBRE_ARMADO)
                Me.Cursor = Cursors.Default
                GoTo SALTO
            End If

            NOMBRE_ARMADO = "tmpQUERY_" & PARAMETRO & "_" & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00") & ".XLS"
            rd.RecordSelectionFormula = SELECCION
            If NOMBRE_DEL_REPORTE_CR = "SPYCS100.rpt" Then
                Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
                Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
                Dim DsCODIGO As New CrystalDecisions.Shared.ParameterDiscreteValue() : DsCODIGO.Value = VALOR01
                Dim DsNOMBRESAPELLIDOS As New CrystalDecisions.Shared.ParameterDiscreteValue() : DsNOMBRESAPELLIDOS.Value = VALOR02
                Dim DsCARGO As New CrystalDecisions.Shared.ParameterDiscreteValue() : DsCARGO.Value = VALOR03
                Dim DsF1 As New CrystalDecisions.Shared.ParameterDiscreteValue() : DsF1.Value = VALOR04
                Dim DsF2 As New CrystalDecisions.Shared.ParameterDiscreteValue() : DsF2.Value = VALOR05
                Dim DsUBICACION As New CrystalDecisions.Shared.ParameterDiscreteValue() : DsUBICACION.Value = VALOR06
                RpDatos.Add(V01) : rd.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
                RpDatos.Add(DsCODIGO) : rd.DataDefinition.ParameterFields("pCODIGO").ApplyCurrentValues(RpDatos)
                RpDatos.Add(DsNOMBRESAPELLIDOS) : rd.DataDefinition.ParameterFields("pNOMBRESAPELLIDOS").ApplyCurrentValues(RpDatos)
                RpDatos.Add(DsCARGO) : rd.DataDefinition.ParameterFields("pCARGO").ApplyCurrentValues(RpDatos)
                RpDatos.Add(DsF1) : rd.DataDefinition.ParameterFields("pF1").ApplyCurrentValues(RpDatos)
                RpDatos.Add(DsF2) : rd.DataDefinition.ParameterFields("pF2").ApplyCurrentValues(RpDatos)
                RpDatos.Add(DsUBICACION) : rd.DataDefinition.ParameterFields("pUBICACION").ApplyCurrentValues(RpDatos)
            End If
            If NOMBRE_DEL_REPORTE_CR = "SPYCS102.rpt" Then
                Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
                Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
                Dim DsF1 As New CrystalDecisions.Shared.ParameterDiscreteValue() : DsF1.Value = VALOR04
                Dim DsF2 As New CrystalDecisions.Shared.ParameterDiscreteValue() : DsF2.Value = VALOR05
                RpDatos.Add(V01) : rd.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
                RpDatos.Add(DsF1) : rd.DataDefinition.ParameterFields("pF1").ApplyCurrentValues(RpDatos)
                RpDatos.Add(DsF2) : rd.DataDefinition.ParameterFields("pF2").ApplyCurrentValues(RpDatos)
            End If

            Dim objExcelOptions As ExcelFormatOptions = New ExcelFormatOptions
            objExcelOptions.ExcelUseConstantColumnWidth = False
            rd.ExportOptions.FormatOptions = objExcelOptions
            Dim strExportFile As String = "C:\WINDOWS\TEMP\" & NOMBRE_ARMADO
            rd.ExportOptions.ExportDestinationType = ExportDestinationType.DiskFile
            rd.ExportOptions.ExportFormatType = ExportFormatType.Excel
            objExcelOptions.ExcelUseConstantColumnWidth = False
            rd.ExportOptions.FormatOptions = objExcelOptions
            Dim objOptions As DiskFileDestinationOptions = New DiskFileDestinationOptions
            objOptions.DiskFileName = strExportFile
            rd.ExportOptions.DestinationOptions = objOptions
            rd.Export()
            objOptions = Nothing
            rd = Nothing
            Process.Start("Excel.exe", "C:\WINDOWS\TEMP\" & NOMBRE_ARMADO)
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
        '////////////////////////////////////////////////////////////////////////////////////////////////////
        '////////////////////////////////////////////////////////////////////////////////////////////////////
        If Me.ComboBox100.Text = "PDF" Then
            Dim cryRpt As New ReportDocument
            Dim PathRPT As String = "\\Hmil-spc-apprd\e$\v3 SPYC\Reportes\"
            If NOMBRE_DEL_REPORTE_CR = "SPYCS100.rpt" Then
                cryRpt.Load(PathRPT & NOMBRE_DEL_REPORTE_CR)
                cryRpt.SetDatabaseLogon("sa", "P@$$W0RD")
                Me.Cursor = Cursors.WaitCursor
                Dim NOMBRE_ARMADO As String
                NOMBRE_ARMADO = "tmpQUERY_" & PARAMETRO & "_" & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00") & ".pdf"
                Dim CrExportOptions As ExportOptions
                Dim CrDiskFileDestinationOptions As New DiskFileDestinationOptions()
                Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()
                cryRpt.RecordSelectionFormula = SELECCION
                Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
                Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
                Dim DsCODIGO As New CrystalDecisions.Shared.ParameterDiscreteValue() : DsCODIGO.Value = VALOR01
                Dim DsNOMBRESAPELLIDOS As New CrystalDecisions.Shared.ParameterDiscreteValue() : DsNOMBRESAPELLIDOS.Value = VALOR02
                Dim DsCARGO As New CrystalDecisions.Shared.ParameterDiscreteValue() : DsCARGO.Value = VALOR03
                Dim DsF1 As New CrystalDecisions.Shared.ParameterDiscreteValue() : DsF1.Value = VALOR04
                Dim DsF2 As New CrystalDecisions.Shared.ParameterDiscreteValue() : DsF2.Value = VALOR05
                Dim DsUBICACION As New CrystalDecisions.Shared.ParameterDiscreteValue() : DsUBICACION.Value = VALOR06
                RpDatos.Add(V01) : cryRpt.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
                RpDatos.Add(DsCODIGO) : cryRpt.DataDefinition.ParameterFields("pCODIGO").ApplyCurrentValues(RpDatos)
                RpDatos.Add(DsNOMBRESAPELLIDOS) : cryRpt.DataDefinition.ParameterFields("pNOMBRESAPELLIDOS").ApplyCurrentValues(RpDatos)
                RpDatos.Add(DsCARGO) : cryRpt.DataDefinition.ParameterFields("pCARGO").ApplyCurrentValues(RpDatos)
                RpDatos.Add(DsF1) : cryRpt.DataDefinition.ParameterFields("pF1").ApplyCurrentValues(RpDatos)
                RpDatos.Add(DsF2) : cryRpt.DataDefinition.ParameterFields("pF2").ApplyCurrentValues(RpDatos)
                RpDatos.Add(DsUBICACION) : cryRpt.DataDefinition.ParameterFields("pUBICACION").ApplyCurrentValues(RpDatos)
                CrDiskFileDestinationOptions.DiskFileName = "C:\WINDOWS\TEMP\" & NOMBRE_ARMADO
                CrExportOptions = cryRpt.ExportOptions
                With CrExportOptions
                    .ExportDestinationType = ExportDestinationType.DiskFile
                    .ExportFormatType = ExportFormatType.PortableDocFormat
                    .DestinationOptions = CrDiskFileDestinationOptions
                    .FormatOptions = CrFormatTypeOptions
                End With
                cryRpt.Export()
                Process.Start("C:\WINDOWS\TEMP\" & NOMBRE_ARMADO)
            Else
                If NOMBRE_DEL_REPORTE_CR = "SPYCS102.rpt" Then
                    cryRpt.Load(PathRPT & NOMBRE_DEL_REPORTE_CR)
                    cryRpt.SetDatabaseLogon("sa", "P@$$W0RD")
                    Me.Cursor = Cursors.WaitCursor
                    Dim NOMBRE_ARMADO As String
                    NOMBRE_ARMADO = "tmpQUERY_" & PARAMETRO & "_" & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00") & ".pdf"
                    Dim CrExportOptions As ExportOptions
                    Dim CrDiskFileDestinationOptions As New DiskFileDestinationOptions()
                    Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()
                    cryRpt.RecordSelectionFormula = SELECCION
                    Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
                    VALOR04 = Mid(Me.DateTimePicker1.Value, 1, 10)
                    VALOR05 = Mid(Me.DateTimePicker2.Value, 1, 10)
                    Dim DsF1 As New CrystalDecisions.Shared.ParameterDiscreteValue() : DsF1.Value = VALOR04
                    Dim DsF2 As New CrystalDecisions.Shared.ParameterDiscreteValue() : DsF2.Value = VALOR05
                    Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
                    RpDatos.Add(V01) : cryRpt.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
                    RpDatos.Add(DsF1) : cryRpt.DataDefinition.ParameterFields("pF1").ApplyCurrentValues(RpDatos)
                    RpDatos.Add(DsF2) : cryRpt.DataDefinition.ParameterFields("pF2").ApplyCurrentValues(RpDatos)
                    CrDiskFileDestinationOptions.DiskFileName = "C:\WINDOWS\TEMP\" & NOMBRE_ARMADO
                    CrExportOptions = cryRpt.ExportOptions
                    With CrExportOptions
                        .ExportDestinationType = ExportDestinationType.DiskFile
                        .ExportFormatType = ExportFormatType.PortableDocFormat
                        .DestinationOptions = CrDiskFileDestinationOptions
                        .FormatOptions = CrFormatTypeOptions
                    End With
                    cryRpt.Export()
                    Process.Start("C:\WINDOWS\TEMP\" & NOMBRE_ARMADO)
                Else
                    If NOMBRE_DEL_REPORTE_CR = "SPYCS103.rpt" Then
                        cryRpt.Load(PathRPT & NOMBRE_DEL_REPORTE_CR)
                        cryRpt.SetDatabaseLogon("sa", "P@$$W0RD")
                        Me.Cursor = Cursors.WaitCursor
                        Dim NOMBRE_ARMADO As String
                        NOMBRE_ARMADO = "tmpQUERY_" & PARAMETRO & "_" & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00") & ".pdf"
                        Dim CrExportOptions As ExportOptions
                        Dim CrDiskFileDestinationOptions As New DiskFileDestinationOptions()
                        Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()
                        cryRpt.RecordSelectionFormula = SELECCION
                        Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
                        CrDiskFileDestinationOptions.DiskFileName = "C:\WINDOWS\TEMP\" & NOMBRE_ARMADO
                        CrExportOptions = cryRpt.ExportOptions
                        With CrExportOptions
                            .ExportDestinationType = ExportDestinationType.DiskFile
                            .ExportFormatType = ExportFormatType.PortableDocFormat
                            .DestinationOptions = CrDiskFileDestinationOptions
                            .FormatOptions = CrFormatTypeOptions
                        End With
                        cryRpt.Export()
                        Process.Start("C:\WINDOWS\TEMP\" & NOMBRE_ARMADO)
                    Else
                        cryRpt.Load(PathRPT & NOMBRE_DEL_REPORTE_CR)
                        cryRpt.SetDatabaseLogon("sa", "P@$$W0RD")
                        Me.Cursor = Cursors.WaitCursor
                        Dim NOMBRE_ARMADO As String
                        NOMBRE_ARMADO = "tmpQUERY_" & PARAMETRO & "_" & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00") & ".pdf"
                        Dim CrExportOptions As ExportOptions
                        Dim CrDiskFileDestinationOptions As New DiskFileDestinationOptions()
                        Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()
                        cryRpt.RecordSelectionFormula = SELECCION
                        Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
                        Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
                        RpDatos.Add(V01) : cryRpt.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
                        CrDiskFileDestinationOptions.DiskFileName = "C:\WINDOWS\TEMP\" & NOMBRE_ARMADO
                        CrExportOptions = cryRpt.ExportOptions
                        With CrExportOptions
                            .ExportDestinationType = ExportDestinationType.DiskFile
                            .ExportFormatType = ExportFormatType.PortableDocFormat
                            .DestinationOptions = CrDiskFileDestinationOptions
                            .FormatOptions = CrFormatTypeOptions
                        End With
                        cryRpt.Export()
                        Process.Start("C:\WINDOWS\TEMP\" & NOMBRE_ARMADO)
                    End If
                End If
            End If
            Me.Cursor = Cursors.Default
            GoTo SALTO
        End If
SALTO:
    End Sub
    Private Sub ComboBox14_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox14.SelectedIndexChanged
        Call CARGAR_COMBO_CAT_DE_SUB_TIPO_DE_MOVIMIENTO_REAL_DE_EXP()
        Call CARGAR_COMBO_CAT_DE_SUB_TIPO_DE_MOVIMIENTO_LEGAL_DE_EXP()
    End Sub
    Private Sub ComboBox17_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox17.SelectedIndexChanged
        Call CARGAR_COMBO_SOL_CAT_DE_RESULTADOS()
    End Sub
    Private Sub ComboBox27_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox27.SelectedIndexChanged
        Call CARGAR_COMBO_CAT_DE_CONDECORACIONES()
    End Sub
    Private Sub ComboBox29_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox29.SelectedIndexChanged
        Call CARGAR_COMBO_CAT_DE_SITUACION_DE_PERSONAL()
    End Sub
    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        If Me.ComboBox27.SelectedValue <> 0 Then
            Me.Lista_Principal.Items.Add("Clasificacion de la Condecoracion: " & Me.ComboBox27.Text)
            If Me.Lista138.Items.Count = 0 Then
                Me.Lista138.Items.Add("({CAT_DE_CLASIFICACION_DE_CONECORACIONES.ID_C_CONDE} = " & Me.ComboBox27.SelectedValue & ")")
            Else
                Me.Lista138.Items.Add(" OR ")
                Me.Lista138.Items.Add("({CAT_DE_CLASIFICACION_DE_CONECORACIONES.ID_C_CONDE} = " & Me.ComboBox27.SelectedValue & ")")
            End If
        Else
            MsgBox("La Seleccion para la Clasificacion de la Condecoracion del Empleado no es válida", vbInformation, "Mensaje del Sistema")
            Me.ComboBox27.Focus()
            Exit Sub
        End If
    End Sub
    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        If Me.ComboBox28.SelectedValue <> 0 Then
            Me.Lista_Principal.Items.Add("Tipo de Condecoracion: " & Me.ComboBox28.Text)
            If Me.Lista139.Items.Count = 0 Then
                Me.Lista139.Items.Add("({MAESTRO_DE_CONDECORACIONES.ID_CONDE} = " & Me.ComboBox28.SelectedValue & ")")
            Else
                Me.Lista139.Items.Add(" OR ")
                Me.Lista139.Items.Add("({MAESTRO_DE_CONDECORACIONES.ID_CONDE} = " & Me.ComboBox28.SelectedValue & ")")
            End If
        Else
            MsgBox("La Seleccion para el Tipo de Condecoracion del Empleado no es válida", vbInformation, "Mensaje del Sistema")
            Me.ComboBox28.Focus()
            Exit Sub
        End If
    End Sub
    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click         'TIPO DE SITUACION
        If Me.ComboBox29.SelectedValue <> 0 Then
            Me.Lista_Principal.Items.Add("Situacion: " & Me.ComboBox29.Text)
            If CODIGO_REPORTE = "S0007" Or CODIGO_REPORTE = "S0008" Or CODIGO_REPORTE = "S0009" Or CODIGO_REPORTE = "S0010" Or CODIGO_REPORTE = "S0011" Or CODIGO_REPORTE = "S0012" Or CODIGO_REPORTE = "S0013" Then
                If Me.Lista141.Items.Count <> 0 Then
                    Me.Lista141.Items.Add(" OR ")
                End If
                Me.Lista141.Items.Add("({MAESTRO_DE_SITUACION_DE_PERSONAL.ID_T_SIT_P} = " & Me.ComboBox29.SelectedValue & ")")
                GoTo SALIR
            End If
        Else
            MsgBox("La Seleccion para el Tipo de Situación no es válida", vbInformation, "Mensaje del Sistema")
            Me.ComboBox29.Focus()
            Exit Sub
        End If
SALIR:
    End Sub
    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click         'SITUACION
        If Me.ComboBox30.SelectedValue <> 0 Then
            Me.Lista_Principal.Items.Add("Tipo de Situacion: " & Me.ComboBox30.Text)
            If CODIGO_REPORTE = "S0007" Or CODIGO_REPORTE = "S0008" Or CODIGO_REPORTE = "S0009" Or CODIGO_REPORTE = "S0010" Or CODIGO_REPORTE = "S0011" Or CODIGO_REPORTE = "S0012" Or CODIGO_REPORTE = "S0013" Then
                If Me.Lista142.Items.Count <> 0 Then
                    Me.Lista142.Items.Add(" OR ")
                End If
                Me.Lista142.Items.Add("({MAESTRO_DE_SITUACION_DE_PERSONAL.ID_SIT_P} = " & Me.ComboBox30.SelectedValue & ")")
                GoTo SALIR
            End If
        Else
            MsgBox("La Seleccion para la Situación no es válida", vbInformation, "Mensaje del Sistema")
            Me.ComboBox30.Focus()
            Exit Sub
        End If
SALIR:
    End Sub
    Private Sub Button29_Click(sender As Object, e As EventArgs) Handles Button29.Click         'FECHAS SITUACION
        If Mid(Me.DateTimePicker3.Value, 1, 10) = Mid(Me.DateTimePicker4.Value, 1, 10) Then
            Me.Lista_Principal.Items.Add("Fecha de Situacion: " & Mid(Me.DateTimePicker3.Value, 1, 10))
        Else
            Me.Lista_Principal.Items.Add("Fecha de Situacion: Del " & Mid(Me.DateTimePicker3.Value, 1, 10) & " al " & Mid(Me.DateTimePicker4.Value, 1, 10))
        End If
        D1 = Mid(Me.DateTimePicker3.Value, 1, 2)
        M1 = Mid(Me.DateTimePicker3.Value, 4, 2)
        A1 = Mid(Me.DateTimePicker3.Value, 7, 4)
        D2 = Mid(Me.DateTimePicker4.Value, 1, 2)
        M2 = Mid(Me.DateTimePicker4.Value, 4, 2)
        A2 = Mid(Me.DateTimePicker4.Value, 7, 4)
        If CODIGO_REPORTE = "S0007" Or CODIGO_REPORTE = "S0008" Or CODIGO_REPORTE = "S0009" Or CODIGO_REPORTE = "S0010" Or CODIGO_REPORTE = "S0011" Or CODIGO_REPORTE = "S0012" Or CODIGO_REPORTE = "S0013" Then
            If Me.Lista141.Items.Count <> 0 Then
                Me.Lista141.Items.Add(" OR ")
            End If
            Me.Lista141.Items.Add("({MAESTRO_DE_SITUACION_DE_PERSONAL.DIA} >= CDateTime (" & A1 & ", " & M1 & ", " & D1 & ", 00, 00, 00) AND {MAESTRO_DE_SITUACION_DE_PERSONAL.DIA} <= CDateTime (" & A2 & ", " & M2 & ", " & D2 & ", 00, 00, 00))")
        End If
    End Sub
    Private Sub TextBox1_DoubleClick(sender As Object, e As EventArgs) Handles TextBox1.DoubleClick
        Frm_Seleccion_Consultas_y_Reportes_Buscar.ShowDialog()
        If CERRAR = False Then
            Me.TextBox1.Text = bCODIGO
            Me.TextBox1.SelectAll()
            Me.TextBox1.Focus()
        End If
    End Sub
End Class
