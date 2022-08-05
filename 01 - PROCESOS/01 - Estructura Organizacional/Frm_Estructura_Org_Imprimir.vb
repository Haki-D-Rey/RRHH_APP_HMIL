Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient

Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class Frm_Estructura_Org_Imprimir
    Private Sub Frm_Estructura_Org_Imprimir_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.ComboBox1.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Estructura_Org_Imprimir_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.CheckBox1.Checked = False
        Call CARGAR_COMBO_CAT_DE_FORMATOS_DE_SALIDA()
        Me.ComboBox1.Focus()
    End Sub
    Private Sub Button03_Click(sender As Object, e As EventArgs) Handles Button03.Click
        Me.ComboBox1.Focus()
        CERRAR = True
        Me.Close()
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
                With Me.ComboBox1
                    .DataSource = MiDataTable
                    .DisplayMember = "DESCRIPCION"
                    .ValueMember = "ID_F_S"
                End With
            Else
                Me.ComboBox1.DataSource = Nothing
            End If
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
    Dim vN11, vN22, vN33, vN44, vN55, vN66, vN77, vN88 As Integer
    Dim vN1, vN2, vN3, vN4, vN5, vN6, vN7, vN8 As String
    Private Sub Button01_Click(sender As Object, e As EventArgs) Handles Button01.Click
        'NIVEL 1
        If Frm_Estructura_Org.ComboBox1.SelectedValue = 0 Then
            vN11 = Nothing
        Else
            vN11 = Frm_Estructura_Org.ComboBox1.SelectedValue
        End If
        'NIVEL 2
        If Frm_Estructura_Org.ComboBox2.SelectedValue = 0 Then
            vN22 = Nothing
        Else
            vN22 = Frm_Estructura_Org.ComboBox2.SelectedValue
        End If
        'NIVEL 3
        If Frm_Estructura_Org.ComboBox3.SelectedValue = 0 Then
            vN33 = Nothing
        Else
            vN33 = Frm_Estructura_Org.ComboBox3.SelectedValue
        End If
        'NIVEL 4
        If Frm_Estructura_Org.ComboBox4.SelectedValue = 0 Then
            vN44 = Nothing
        Else
            vN44 = Frm_Estructura_Org.ComboBox4.SelectedValue
        End If
        'NIVEL 5
        If Frm_Estructura_Org.ComboBox5.SelectedValue = 0 Then
            vN55 = Nothing
        Else
            vN55 = Frm_Estructura_Org.ComboBox5.SelectedValue
        End If
        'NIVEL 6
        If Frm_Estructura_Org.ComboBox6.SelectedValue = 0 Then
            vN66 = Nothing
        Else
            vN66 = Frm_Estructura_Org.ComboBox6.SelectedValue
        End If
        'NIVEL 7
        If Frm_Estructura_Org.ComboBox7.SelectedValue = 0 Then
            vN77 = Nothing
        Else
            vN77 = Frm_Estructura_Org.ComboBox7.SelectedValue
        End If
        'NIVEL 8
        If Frm_Estructura_Org.ComboBox8.SelectedValue = 0 Then
            vN88 = Nothing
        Else
            vN88 = Frm_Estructura_Org.ComboBox8.SelectedValue
        End If

        If vN11 = 0 Then : vN1 = 0 : Else : vN1 = Frm_Estructura_Org.ComboBox1.SelectedValue : End If
        If vN22 = 0 Then : vN2 = 0 : Else : vN2 = Frm_Estructura_Org.ComboBox2.SelectedValue : End If
        If vN33 = 0 Then : vN3 = 0 : Else : vN3 = Frm_Estructura_Org.ComboBox3.SelectedValue : End If
        If vN44 = 0 Then : vN4 = 0 : Else : vN4 = Frm_Estructura_Org.ComboBox4.SelectedValue : End If
        If vN55 = 0 Then : vN5 = 0 : Else : vN5 = Frm_Estructura_Org.ComboBox5.SelectedValue : End If
        If vN66 = 0 Then : vN6 = 0 : Else : vN6 = Frm_Estructura_Org.ComboBox6.SelectedValue : End If
        If vN77 = 0 Then : vN7 = 0 : Else : vN7 = Frm_Estructura_Org.ComboBox7.SelectedValue : End If
        If vN88 = 0 Then : vN8 = 0 : Else : vN8 = Frm_Estructura_Org.ComboBox8.SelectedValue : End If
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        If Me.ComboBox1.Text = "CRYSTAL" Then
            SELECCION_PARAMETRO = "*"
            USUARIO_IMPRIME = isuCUENTA
            Frm_Visor_Reportes.CrystalR.ShowExportButton = True
            Frm_Visor_Reportes.CrystalR.ShowPrintButton = True
            If isuID_TIPO_USUARIO = 5 Then
                If Me.CheckBox1.Checked = False Then
                    If vN2 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND ({MAESTRO_DE_PERSONAS.ID_ADMIN} = " & isuUSUARIO_R & ")"
                        GoTo SALTO
                    End If
                    If vN3 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND ({MAESTRO_DE_PERSONAS.ID_ADMIN} = " & isuUSUARIO_R & ")"
                        GoTo SALTO
                    End If
                    If vN4 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ") AND ({MAESTRO_DE_PERSONAS.ID_ADMIN} = " & isuUSUARIO_R & ")"
                        GoTo SALTO
                    End If
                    If vN5 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N4} = " & vN4 & ") AND ({MAESTRO_DE_PERSONAS.ID_ADMIN} = " & isuUSUARIO_R & ")"
                        GoTo SALTO
                    End If
                    If vN6 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N4} = " & vN4 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N5} = " & vN5 & ") AND ({MAESTRO_DE_PERSONAS.ID_ADMIN} = " & isuUSUARIO_R & ")"
                        GoTo SALTO
                    End If
                    If vN7 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N4} = " & vN4 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N5} = " & vN5 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N6} = " & vN6 & ") AND ({MAESTRO_DE_PERSONAS.ID_ADMIN} = " & isuUSUARIO_R & ")"
                        GoTo SALTO
                    End If
                    If vN8 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N4} = " & vN4 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N5} = " & vN5 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N6} = " & vN6 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N7} = " & vN7 & ") AND ({MAESTRO_DE_PERSONAS.ID_ADMIN} = " & isuUSUARIO_R & ")"
                        GoTo SALTO
                    End If
                    'If vN2 <> 0 And vN3 <> 0 And vN4 <> 0 And vN5 <> 0 And vN6 <> 0 And vN7 <> 0 And vN8 <> 0 Then
                    '    SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N4} = " & vN4 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N5} = " & vN5 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N6} = " & vN6 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N7} = " & vN7 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N8} = " & vN8 & ") AND ({MAESTRO_DE_PERSONAS.ID_ADMIN} = " & isuUSUARIO_R & ")"
                    '    GoTo SALTO
                    'End If
                Else
                    SELECCION = "({MAESTRO_DE_PERSONAS.ID_ADMIN} = " & isuUSUARIO_R & ")"
                End If
            Else
                If Me.CheckBox1.Checked = False Then
                    If vN2 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ")"
                        GoTo SALTO
                    End If
                    If vN3 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ")"
                        GoTo SALTO
                    End If
                    If vN4 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ")"
                        GoTo SALTO
                    End If
                    If vN5 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N4} = " & vN4 & ")"
                        GoTo SALTO
                    End If
                    If vN6 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N4} = " & vN4 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N5} = " & vN5 & ")"
                        GoTo SALTO
                    End If
                    If vN7 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N4} = " & vN4 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N5} = " & vN5 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N6} = " & vN6 & ")"
                        GoTo SALTO
                    End If
                    If vN8 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N4} = " & vN4 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N5} = " & vN5 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N6} = " & vN6 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N7} = " & vN7 & ")"
                        GoTo SALTO
                    End If
                    'If vN2 <> 0 And vN3 <> 0 And vN4 <> 0 And vN5 <> 0 And vN6 <> 0 And vN7 <> 0 And vN8 <> 0 Then
                    '    SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N4} = " & vN4 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N5} = " & vN5 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N6} = " & vN6 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N7} = " & vN7 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N8} = " & vN8 & ")"
                    '    GoTo SALTO
                    'End If
                Else
                    SELECCION = ""
                End If
            End If
SALTO:
                PARAMETRO = 2
                Frm_Visor_Reportes.ShowDialog()
            End If
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        If Me.ComboBox1.Text = "OFFICE (EXCEL\WORD)" Then
            If isuID_TIPO_USUARIO = 5 Then
                If Me.CheckBox1.Checked = False Then
                    If vN2 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND ({MAESTRO_DE_PERSONAS.ID_ADMIN} = " & isuUSUARIO_R & ")"
                        GoTo SALTO1
                    End If
                    If vN3 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND ({MAESTRO_DE_PERSONAS.ID_ADMIN} = " & isuUSUARIO_R & ")"
                        GoTo SALTO1
                    End If
                    If vN4 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ") AND ({MAESTRO_DE_PERSONAS.ID_ADMIN} = " & isuUSUARIO_R & ")"
                        GoTo SALTO1
                    End If
                    If vN5 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N4} = " & vN4 & ") AND ({MAESTRO_DE_PERSONAS.ID_ADMIN} = " & isuUSUARIO_R & ")"
                        GoTo SALTO1
                    End If
                    If vN6 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N4} = " & vN4 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N5} = " & vN5 & ") AND ({MAESTRO_DE_PERSONAS.ID_ADMIN} = " & isuUSUARIO_R & ")"
                        GoTo SALTO1
                    End If
                    If vN7 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N4} = " & vN4 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N5} = " & vN5 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N6} = " & vN6 & ") AND ({MAESTRO_DE_PERSONAS.ID_ADMIN} = " & isuUSUARIO_R & ")"
                        GoTo SALTO1
                    End If
                    If vN8 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N4} = " & vN4 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N5} = " & vN5 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N6} = " & vN6 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N7} = " & vN7 & ") AND ({MAESTRO_DE_PERSONAS.ID_ADMIN} = " & isuUSUARIO_R & ")"
                        GoTo SALTO1
                    End If
                    'If vN8 <> 0 Then
                    '    SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N4} = " & vN4 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N5} = " & vN5 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N6} = " & vN6 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N7} = " & vN7 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N8} = " & vN8 & ") AND ({MAESTRO_DE_PERSONAS.ID_ADMIN} = " & isuUSUARIO_R & ")"
                    '    GoTo SALTO1
                    'End If
                Else
                    SELECCION = "({MAESTRO_DE_PERSONAS.ID_ADMIN} = " & isuUSUARIO_R & ")"
                End If
            Else
                If Me.CheckBox1.Checked = False Then
                    If vN2 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ")"
                        GoTo SALTO1
                    End If
                    If vN3 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ")"
                        GoTo SALTO1
                    End If
                    If vN4 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ")"
                        GoTo SALTO1
                    End If
                    If vN5 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N4} = " & vN4 & ")"
                        GoTo SALTO1
                    End If
                    If vN6 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N4} = " & vN4 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N5} = " & vN5 & ")"
                        GoTo SALTO1
                    End If
                    If vN7 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N4} = " & vN4 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N5} = " & vN5 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N6} = " & vN6 & ")"
                        GoTo SALTO1
                    End If
                    If vN8 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N4} = " & vN4 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N5} = " & vN5 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N6} = " & vN6 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N7} = " & vN7 & ")"
                        GoTo SALTO1
                    End If
                    'If vN2 <> 0 And vN3 <> 0 And vN4 <> 0 And vN5 <> 0 And vN6 <> 0 And vN7 <> 0 And vN8 <> 0 Then
                    '    SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N4} = " & vN4 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N5} = " & vN5 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N6} = " & vN6 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N7} = " & vN7 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N8} = " & vN8 & ")"
                    '    GoTo SALTO1
                    'End If
                Else
                    SELECCION = ""
                End If
            End If
SALTO1:
            Dim rd = New ReportDocument()
            Dim PathRPT As String = "\\Hmil-spc-apprd\e$\v3 SPYC\Excel\"
            rd.Load(PathRPT & "SPYC002x.rpt")
            rd.SetDatabaseLogon("sa", "P@$$W0RD")
            If SELECCION_PARAMETRO = "" Then
                SELECCION_PARAMETRO = "*"
            End If
            On Error GoTo SALIR
            Me.Cursor = Cursors.WaitCursor
            Dim NOMBRE_ARMADO As String
            NOMBRE_ARMADO = "tmpQUERY_" & PARAMETRO & "_" & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00") & ".XLS"
            rd.RecordSelectionFormula = SELECCION
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
            Exit Sub
SALIR:
            MsgBox("Se ha generado un error, posiblemente no se han seleccionado los parametros requeridos para generar el archivo de Excel", vbInformation, "Mensaje del Sistema")
            Me.Cursor = Cursors.Default
        End If
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        If Me.ComboBox1.Text = "PDF" Then
            If isuID_TIPO_USUARIO = 5 Then
                If Me.CheckBox1.Checked = False Then
                    If vN2 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND ({MAESTRO_DE_PERSONAS.ID_ADMIN} = " & isuUSUARIO_R & ")"
                        GoTo SALTO3
                    End If
                    If vN3 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND ({MAESTRO_DE_PERSONAS.ID_ADMIN} = " & isuUSUARIO_R & ")"
                        GoTo SALTO3
                    End If
                    If vN4 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ") AND ({MAESTRO_DE_PERSONAS.ID_ADMIN} = " & isuUSUARIO_R & ")"
                        GoTo SALTO3
                    End If
                    If vN5 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N4} = " & vN4 & ") AND ({MAESTRO_DE_PERSONAS.ID_ADMIN} = " & isuUSUARIO_R & ")"
                        GoTo SALTO3
                    End If
                    If vN6 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N4} = " & vN4 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N5} = " & vN5 & ") AND ({MAESTRO_DE_PERSONAS.ID_ADMIN} = " & isuUSUARIO_R & ")"
                        GoTo SALTO3
                    End If
                    If vN7 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N4} = " & vN4 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N5} = " & vN5 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N6} = " & vN6 & ") AND ({MAESTRO_DE_PERSONAS.ID_ADMIN} = " & isuUSUARIO_R & ")"
                        GoTo SALTO3
                    End If
                    If vN8 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N4} = " & vN4 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N5} = " & vN5 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N6} = " & vN6 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N7} = " & vN7 & ") AND ({MAESTRO_DE_PERSONAS.ID_ADMIN} = " & isuUSUARIO_R & ")"
                        GoTo SALTO3
                    End If
                    'If vN2 <> 0 And vN3 <> 0 And vN4 <> 0 And vN5 <> 0 And vN6 <> 0 And vN7 <> 0 And vN8 <> 0 Then
                    '    SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N4} = " & vN4 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N5} = " & vN5 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N6} = " & vN6 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N7} = " & vN7 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N8} = " & vN8 & ") AND ({MAESTRO_DE_PERSONAS.ID_ADMIN} = " & isuUSUARIO_R & ")"
                    '    GoTo SALTO3
                    'End If
                Else
                    SELECCION = "({MAESTRO_DE_PERSONAS.ID_ADMIN} = " & isuUSUARIO_R & ")"
                End If
            Else
                If Me.CheckBox1.Checked = False Then
                    If vN2 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ")"
                        GoTo SALTO3
                    End If
                    If vN3 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ")"
                        GoTo SALTO3
                    End If
                    If vN4 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ")"
                        GoTo SALTO3
                    End If
                    If vN5 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N4} = " & vN4 & ")"
                        GoTo SALTO3
                    End If
                    If vN6 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N4} = " & vN4 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N5} = " & vN5 & ")"
                        GoTo SALTO3
                    End If
                    If vN7 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N4} = " & vN4 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N5} = " & vN5 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N6} = " & vN6 & ")"
                        GoTo SALTO3
                    End If
                    If vN8 = 0 Then
                        SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N4} = " & vN4 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N5} = " & vN5 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N6} = " & vN6 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N7} = " & vN7 & ")"
                        GoTo SALTO3
                    End If
                    'If vN2 <> 0 And vN3 <> 0 And vN4 <> 0 And vN5 <> 0 And vN6 <> 0 And vN7 <> 0 And vN8 <> 0 Then
                    '    SELECCION = "({MAESTRO_DE_CARGOS.ID_N1} = " & vN1 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N2} = " & vN2 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N3} = " & vN3 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N4} = " & vN4 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N5} = " & vN5 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N6} = " & vN6 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N7} = " & vN7 & ") AND " & "({MAESTRO_DE_CARGOS.ID_N8} = " & vN8 & ")"
                    '    GoTo SALTO3
                    'End If
                Else
                    SELECCION = ""
                End If
            End If
SALTO3:
                Dim cryRpt As New ReportDocument
                Dim PathRPT As String = "\\Hmil-spc-apprd\e$\v3 SPYC\Reportes\"
                cryRpt.Load(PathRPT & "SPYC002.rpt")
                cryRpt.SetDatabaseLogon("sa", "P@$$W0RD")
                If SELECCION_PARAMETRO = "" Then
                    SELECCION_PARAMETRO = "*"
                End If
                On Error GoTo SALIR
                Me.Cursor = Cursors.WaitCursor
                Dim NOMBRE_ARMADO As String
                NOMBRE_ARMADO = "tmpQUERY_" & PARAMETRO & "_" & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00") & ".pdf"
                Dim CrExportOptions As ExportOptions
                Dim CrDiskFileDestinationOptions As New DiskFileDestinationOptions()
                Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()
                cryRpt.RecordSelectionFormula = SELECCION
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
                Me.Cursor = Cursors.Default
                Exit Sub
                Me.Cursor = Cursors.Default
            End If
    End Sub
End Class