Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Formato_Vacaciones
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        'VALOR01 = Me.TextBox1.Text                          'CODIGO
        VALOR01 = Me.TextBox2.Text                          'NOMBRES Y APELLIDOS
        VALOR02 = Me.TextBox3.Text                          'CARGO
        VALOR03 = Me.TextBox4.Text                          'UBICACION
        VALOR04 = Me.TextBox5.Text                          'CANTIDAD DE DIAS
        VALOR05 = Me.TextBox6.Text                          'CANTIDAD DE HORAS       
        If Me.RadioButton1.Checked = True Then : VALOR06 = "X" : Else : VALOR06 = "" : End If
        If Me.RadioButton2.Checked = True Then : VALOR07 = "X" : Else : VALOR07 = "" : End If
        If Me.RadioButton3.Checked = True Then : VALOR08 = "X" : Else : VALOR08 = "" : End If
        If Me.RadioButton4.Checked = True Then : VALOR09 = "X" : Else : VALOR09 = "" : End If
        If Me.RadioButton5.Checked = True Then : VALOR15 = "X" : Else : VALOR15 = "" : End If

        If Me.DateTimePicker1.Checked = True Then : VALOR10 = Mid(Me.DateTimePicker1.Value, 1, 10) : Else : VALOR10 = "" : End If
        If Me.DateTimePicker2.Checked = True Then : VALOR11 = Mid(Me.DateTimePicker2.Value, 1, 10) : Else : VALOR11 = "" : End If

        VALOR12 = Me.TextBox9.Text                          'OBSERVACIONES LINEA 1
        VALOR13 = Me.TextBox8.Text                          'OBSERVACIONES LINEA 2
        VALOR14 = Me.TextBox7.Text                          'OBSERVACIONES LINEA 3

        SELECCION = ""
        SELECCION_PARAMETRO = "*"
        USUARIO_IMPRIME = isuCUENTA
        NOMBRE_DEL_REPORTE_CR = "SPYCS103.rpt"
        PARAMETRO = 139
        Frm_Visor_Reportes.ShowDialog()
    End Sub
    Private Sub Frm_Formato_Vacaciones_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Me.TextBox6.Text = ""
        Me.RadioButton1.Checked = False
        Me.RadioButton2.Checked = False
        Me.RadioButton3.Checked = False
        Me.RadioButton4.Checked = False
        Me.RadioButton5.Checked = False
        Call Clases.BUSCAR_FECHA_Y_HORA()
        Me.DateTimePicker1.Value = Mid(FECHA_SERVIDOR, 1, 10)
        Me.DateTimePicker2.Value = Mid(FECHA_SERVIDOR, 1, 10)
        Me.TextBox9.Text = ""
        Me.TextBox8.Text = ""
        Me.TextBox7.Text = ""
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Me.Close()
    End Sub
    Private Sub Frm_Formato_Vacaciones_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.Close()
        End If
    End Sub
    Dim id_EMPLEADO As Integer
    Dim cEMPLEADO As String
    Dim DATO_1, DATO_2 As Boolean
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            cEMPLEADO = CInt(Val(TextBox1.Text))
            TextBox1.Text = Format(CInt(cEMPLEADO), "0000000000")
            If Len(Me.TextBox1.Text) = 0 Or (Me.TextBox1.Text = "0000000000") Then
                MsgBox("El código que se digitó no es válido", vbInformation, "Mensaje del Sistema")
                Me.TextBox1.Focus()
                Exit Sub
            Else
                CADENAsql = "SELECT ID_M_P, CODIGO, APELLIDOS, NOMBRES, N_CEDULA, FEC_ING_PAME, ID_ESTADO_P, ID_N_PROFESIONAL, ID_T_SEXO, PREFIJO FROM [dbo].[VISTA MAESTRO DE PERSONAS] WHERE CODIGO = '" & Me.TextBox1.Text & "' AND ID_ESTADO_P = 1"
                Call BUSCAR_DATO_1()
                If DATO_1 = True Then
                    Call BUSCAR_DATO_2_ACTIVOS()
                    e.Handled = True
                    SendKeys.Send("{TAB}")
                Else
                    MsgBox("El código que se digitó no es válido", vbInformation, "Mensaje del Sistema")
                    Me.TextBox1.SelectAll()
                    Me.TextBox1.Focus()
                    Exit Sub
                End If
            End If
        End If
    End Sub
    Private Sub BUSCAR_DATO_2_ACTIVOS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT CODIGO, D_N1, D_N2, D_N3, D_N4, D_N5, D_N6, D_N7, D_N8, D_CARGO_ES, ID_SITUACION FROM [dbo].[VISTA MAESTRO DE CARGOS SIT] WHERE CODIGO = '" & Me.TextBox1.Text & "' AND ID_SITUACION = 1", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                'Me.TextBox5.Text = MiDataTable.Rows(0).Item(1).ToString
                Dim cadenaUBIC As String
                cadenaUBIC = MiDataTable.Rows(0).Item(1).ToString
                If MiDataTable.Rows(0).Item(2).ToString = "" Then
                    GoTo SALIR
                Else
                    cadenaUBIC = cadenaUBIC & "\ " & MiDataTable.Rows(0).Item(2).ToString
                End If
                If MiDataTable.Rows(0).Item(3).ToString = "" Then
                    GoTo SALIR
                Else
                    cadenaUBIC = cadenaUBIC & "\ " & MiDataTable.Rows(0).Item(3).ToString
                End If
                If MiDataTable.Rows(0).Item(4).ToString = "" Then
                    GoTo SALIR
                Else
                    cadenaUBIC = cadenaUBIC & "\ " & MiDataTable.Rows(0).Item(4).ToString
                End If
                If MiDataTable.Rows(0).Item(5).ToString = "" Then
                    GoTo SALIR
                Else
                    cadenaUBIC = cadenaUBIC & "\ " & MiDataTable.Rows(0).Item(5).ToString
                End If
                If MiDataTable.Rows(0).Item(6).ToString = "" Then
                    GoTo SALIR
                Else
                    cadenaUBIC = cadenaUBIC & "\ " & MiDataTable.Rows(0).Item(6).ToString
                End If
                If MiDataTable.Rows(0).Item(7).ToString = "" Then
                    GoTo SALIR
                Else
                    cadenaUBIC = cadenaUBIC & "\ " & MiDataTable.Rows(0).Item(7).ToString
                End If
                If MiDataTable.Rows(0).Item(8).ToString = "" Then
                    GoTo SALIR
                Else
                    cadenaUBIC = cadenaUBIC & "\ " & MiDataTable.Rows(0).Item(8).ToString
                End If
SALIR:
                Me.TextBox4.Text = cadenaUBIC                                      'UBICACION
                Me.TextBox3.Text = MiDataTable.Rows(0).Item(9).ToString
                DATO_2 = True
            Else
                Me.TextBox4.Text = ""
                Me.TextBox3.Text = ""
                DATO_2 = False
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim NP As Integer
    Dim SEX As String
    Private Sub BUSCAR_DATO_1()
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
                id_EMPLEADO = MiDataTable.Rows(0).Item(0).ToString
                Me.TextBox2.Text = MiDataTable.Rows(0).Item(3).ToString & " " & MiDataTable.Rows(0).Item(2).ToString
                'Me.TextBox3.Text = MiDataTable.Rows(0).Item(4).ToString
                'Me.DateTimePicker1.Value = MiDataTable.Rows(0).Item(5).ToString
                'If MiDataTable.Rows(0).Item(8).ToString = 1 Then    'FEMENINO
                'NP = MiDataTable.Rows(0).Item(7).ToString
                'SEX = "la "
                'End If
                'If MiDataTable.Rows(0).Item(8).ToString = 2 Then    'MASCULINO
                'NP = MiDataTable.Rows(0).Item(7).ToString
                'SEX = "el "
                'End If
                'If MiDataTable.Rows(0).Item(8).ToString = 3 Then    '.
                '    NP = MiDataTable.Rows(0).Item(7).ToString
                '    SEX = " "
                'End If
                'Me.TextBox6.Text = MiDataTable.Rows(0).Item(9).ToString
                DATO_1 = True
            Else
                Me.TextBox2.Text = ""
                'Me.TextBox3.Text = ""
                'Me.DateTimePicker1.Value = Now
                NP = 0
                Me.TextBox6.Text = ""
                DATO_1 = False
            End If
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
    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
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
    Private Sub TextBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox9.KeyDown
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
    Private Sub TextBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox7.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
End Class