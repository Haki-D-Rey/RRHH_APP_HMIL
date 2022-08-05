Imports System.ComponentModel
Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_pBACLIST_Actualizar
    Dim F1, F2 As String
    Private Sub Frm_pBACLIST_Actualizar_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Conexion.CBD_RRHH()
        Me.Cursor = Cursors.Default
    End Sub
    'Dim vID_M_P As Integer
    Private Sub Frm_pBACLIST_Actualizar_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Show()
        If Frm_Consulta_Matriz_BAC.CheckBox1.Checked = True Then
            FECHA1 = Mid(Frm_Consulta_Matriz_BAC.DateTimePicker1.Value, 1, 10)
            F1 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA1 + "', 105), 23)"
            FECHA2 = Mid(Frm_Consulta_Matriz_BAC.DateTimePicker2.Value, 1, 10)
            F2 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA2 + "', 105), 23)"
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE CARGOS SIT 2] WHERE FEC_ING_PAME BETWEEN " & F1 & " AND " & F2 & " ORDER BY ID_M_C"
        Else
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE CARGOS SIT 2] ORDER BY ID_M_C"
        End If
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
                For X = 0 To MiDataTable.Rows.Count - 1
                    Application.DoEvents()
                    Me.Cursor = Cursors.WaitCursor
                    vID_M_P = MiDataTable.Rows(X).Item(39).ToString         '
                    vNA = Trim(MiDataTable.Rows(X).Item(41).ToString)       '
                    vN = Trim(MiDataTable.Rows(X).Item(42).ToString)        '
                    vA = Trim(MiDataTable.Rows(X).Item(43).ToString)        '
                    vREFERENCIA = MiDataTable.Rows(X).Item(40).ToString     '
                    Call DIVIDIR_NA()
                    vOCUPACION = MiDataTable.Rows(X).Item(32).ToString      '
                    bacN1 = MiDataTable.Rows(X).Item(4).ToString            '
                    bacN2 = MiDataTable.Rows(X).Item(7).ToString            '   
                    bacN3 = MiDataTable.Rows(X).Item(10).ToString           '
                    bacN4 = MiDataTable.Rows(X).Item(13).ToString           '
                    bacN5 = MiDataTable.Rows(X).Item(16).ToString           '
                    bacN6 = MiDataTable.Rows(X).Item(19).ToString           '
                    bacN7 = MiDataTable.Rows(X).Item(22).ToString           '
                    bacN8 = MiDataTable.Rows(X).Item(25).ToString           '
                    vGRADO = MiDataTable.Rows(X).Item(36).ToString          '
                    vMILITAR = MiDataTable.Rows(X).Item(48).ToString        '
                    vPAME = MiDataTable.Rows(X).Item(49).ToString           '
                    vJEFE = MiDataTable.Rows(X).Item(46).ToString           '
                    vORDEN_ESTRUCTURA = MiDataTable.Rows(X).Item(1).ToString    '
                    Call BUSCAR_DATOS_DE_PERSONA()
                    Call BUSCAR_DATOS_CONYUGUE()
                    Call BUSCAR_DATOS_HIJOS()
                    Call BUSCAR_DATOS_INGRESO()
                    Call INGRESAR_REGISTRO()
                    Me.Label1.Text = "Empleado: " & vNA
                    Me.Label1.Refresh()
                    Me.Label2.Text = " Actualizando: " & X & " DE " & MiDataTable.Rows.Count & " ... "
                    Me.Label2.Refresh()
                    Me.Cursor = Cursors.Default
                Next
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
        Me.Close()
    End Sub
    Private Sub BUSCAR_DATOS_INGRESO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable1 As New DataTable
        Dim MiDataAdapter1 As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter1 = New SqlDataAdapter("SELECT * FROM [dbo].[VISTA MAESTRO DE NOMINAS Y SALARIOS]  WHERE ID_M_P = " & vID_M_P & " ORDER BY F_MOVIMIENTO_N DESC", Conexion.CxRRHH)
            MiDataTable1.Clear()
            MiDataAdapter1.Fill(MiDataTable1)
            If MiDataTable1.Rows.Count > 0 Then
                If MiDataTable1.Rows(0).Item(8).ToString = "0" Or MiDataTable1.Rows(0).Item(8).ToString = "" Then
                    vINGRESO = ""
                Else
                    vINGRESO = MiDataTable1.Rows(0).Item(8).ToString
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
    Private Sub BUSCAR_DATOS_HIJOS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable1 As New DataTable
        Dim MiDataAdapter1 As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter1 = New SqlDataAdapter("SELECT * FROM [VISTA MAESTRO DE NUCLEO FAMILIAR] WHERE ID_M_P = " & vID_M_P & " AND ID_PARENTELA > 4 ORDER BY ID_M_P", Conexion.CxRRHH)
            MiDataTable1.Clear()
            MiDataAdapter1.Fill(MiDataTable1)
            If MiDataTable1.Rows.Count > 0 Then
                vCANT_DEPENDIENTES = Val(MiDataTable1.Rows.Count)
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub BUSCAR_DATOS_CONYUGUE()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable1 As New DataTable
        Dim MiDataAdapter1 As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter1 = New SqlDataAdapter("SELECT * FROM [VISTA MAESTRO DE NUCLEO FAMILIAR] WHERE ID_M_P = " & vID_M_P & " AND ID_PARENTELA = 4 ORDER BY ID_M_P", Conexion.CxRRHH)
            MiDataTable1.Clear()
            MiDataAdapter1.Fill(MiDataTable1)
            If MiDataTable1.Rows.Count > 0 Then
                vCONYUGE = MiDataTable1.Rows(0).Item(4).ToString
                If MiDataTable1.Rows(0).Item(12).ToString = "." Or MiDataTable1.Rows(0).Item(12).ToString = "0" Or Len(MiDataTable1.Rows(0).Item(12).ToString) <> 8 Then
                    vCELULAR_CONYUGE = ""
                Else
                    vCELULAR_CONYUGE = MiDataTable1.Rows(0).Item(12).ToString
                End If
                If MiDataTable1.Rows(0).Item(13).ToString = "." Or MiDataTable1.Rows(0).Item(13).ToString = "0" Or Len(MiDataTable1.Rows(0).Item(13).ToString) <> 8 Then
                    vCELULAR_CONYUGE = vCELULAR_CONYUGE
                Else
                    If vCELULAR_CONYUGE <> "" Then
                        vCELULAR_CONYUGE = vCELULAR_CONYUGE & ", " & MiDataTable1.Rows(0).Item(13).ToString
                    Else
                        vCELULAR_CONYUGE = MiDataTable1.Rows(0).Item(13).ToString
                    End If
                End If
            Else
                vCONYUGE = ""
                vCELULAR_CONYUGE = ""
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub BUSCAR_DATOS_DE_PERSONA()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable1 As New DataTable
        Dim MiDataAdapter1 As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter1 = New SqlDataAdapter("SELECT * FROM [VISTA MAESTRO DE PERSONAS] WHERE ID_M_P = " & vID_M_P & " ORDER BY ID_M_P", Conexion.CxRRHH)
            MiDataTable1.Clear()
            MiDataAdapter1.Fill(MiDataTable1)
            If MiDataTable1.Rows.Count > 0 Then
                For X = 0 To MiDataTable1.Rows.Count - 1
                    Application.DoEvents()
                    Me.Cursor = Cursors.WaitCursor
                    If Len(MiDataTable1.Rows(X).Item(9).ToString) < 3 Then
                        vSEXO = ""
                    Else
                        vSEXO = MiDataTable1.Rows(X).Item(9).ToString
                    End If
                    If Len(MiDataTable1.Rows(X).Item(50).ToString) < 2 Then
                        vESTADO_CIVIL = ""
                    Else
                        vESTADO_CIVIL = MiDataTable1.Rows(X).Item(50).ToString
                    End If
                    If Len(MiDataTable1.Rows(X).Item(34).ToString) < 5 Then
                        vPROFESION_OFICIO = ""
                    Else
                        vPROFESION_OFICIO = MiDataTable1.Rows(X).Item(34).ToString
                    End If
                    Call BUSCAR_OFICIO()
                    If vNOMBRE_ESTUDIO <> "" Then
                        If vPROFESION_OFICIO <> "" Then
                            vPROFESION_OFICIO = vPROFESION_OFICIO & "\ " & vNOMBRE_ESTUDIO
                        Else
                            vPROFESION_OFICIO = vNOMBRE_ESTUDIO
                        End If
                    End If
                    vFECHA_INGRESO = Mid(MiDataTable1.Rows(X).Item(2).ToString, 1, 10)
                    vBENEFICIARIO = ""
                    vFECHA_NACIMIENTO = Mid(MiDataTable1.Rows(X).Item(1).ToString, 1, 10)
                    If MiDataTable1.Rows(X).Item(36).ToString = "MEXICANA" Then
                        vPAIS_NACIMIENTO = "MEXICO"
                    End If
                    If MiDataTable1.Rows(X).Item(36).ToString = "CUBANA" Then
                        vPAIS_NACIMIENTO = "CUBA"
                    End If
                    If MiDataTable1.Rows(X).Item(36).ToString = "VENEZOLANA" Then
                        vPAIS_NACIMIENTO = "VENEZUELA"
                    End If
                    If MiDataTable1.Rows(X).Item(36).ToString = "NICARAGUENSE" Then
                        vPAIS_NACIMIENTO = "NICARAGUA"
                    End If
                    If MiDataTable1.Rows(X).Item(36).ToString = "COLOMBIANA" Then
                        vPAIS_NACIMIENTO = "COLOMBIA"
                    End If
                    If MiDataTable1.Rows(X).Item(36).ToString = "SALVADOREÑA" Then
                        vPAIS_NACIMIENTO = "EL SALVADOR"
                    End If
                    If MiDataTable1.Rows(X).Item(36).ToString = "COSTARRICENSE" Then
                        vPAIS_NACIMIENTO = "COSTA RICA"
                    End If
                    If MiDataTable1.Rows(X).Item(36).ToString = "BELICEÑA" Then
                        vPAIS_NACIMIENTO = "BELICE"
                    End If
                    If MiDataTable1.Rows(X).Item(36).ToString = "HONDUREÑA" Then
                        vPAIS_NACIMIENTO = "HONDURAS"
                    End If
                    If MiDataTable1.Rows(X).Item(36).ToString = "." Then
                        vPAIS_NACIMIENTO = ""
                    End If
                    If Len(MiDataTable1.Rows(X).Item(36).ToString) < 5 Then
                        vNACIONALIDAD = ""
                    Else
                        vNACIONALIDAD = MiDataTable1.Rows(X).Item(36).ToString
                    End If

                    If MiDataTable1.Rows(X).Item(43).ToString = "VENEZOLANA" Then
                        vPAIS_RESIDENCIA = "VENEZUELA"
                    End If
                    If MiDataTable1.Rows(X).Item(43).ToString = "NICARAGUENSE" Then
                        vPAIS_RESIDENCIA = "NICARAGUA"
                    End If
                    If MiDataTable1.Rows(X).Item(43).ToString = "SALVADOREÑA" Then
                        vPAIS_RESIDENCIA = "EL SALVADOR"
                    End If
                    If MiDataTable1.Rows(X).Item(43).ToString = "COSTARRICENSE" Then
                        vPAIS_RESIDENCIA = "COSTA RICA"
                    End If
                    If MiDataTable1.Rows(X).Item(43).ToString = "HONDUREÑA" Then
                        vPAIS_RESIDENCIA = "HONDURA"
                    End If
                    If MiDataTable1.Rows(X).Item(43).ToString = "." Then
                        vPAIS_RESIDENCIA = ""
                    End If
                    If Len(Trim(MiDataTable1.Rows(X).Item(10).ToString)) > 5 Then
                        vDIRECCION_DOMICILIAR = Trim(MiDataTable1.Rows(X).Item(10).ToString)
                    Else
                        vDIRECCION_DOMICILIAR = ""
                    End If
                    If Len(vDIRECCION_DOMICILIAR) < 3 Then
                        vDIRECCION_DOMICILIAR = Trim(MiDataTable1.Rows(X).Item(11).ToString)
                    End If
                    If Len(vDIRECCION_DOMICILIAR) < 3 Then
                        vDIRECCION_DOMICILIAR = ""
                    End If
                    If Len(Trim(MiDataTable1.Rows(X).Item(47).ToString)) < 5 Then
                        vMUNICIPIO_DOMICILIAR = ""
                    Else
                        vMUNICIPIO_DOMICILIAR = Trim(MiDataTable1.Rows(X).Item(47).ToString)
                    End If
                    If Len(Trim(MiDataTable1.Rows(X).Item(45).ToString)) < 5 Then
                        vDEPARTAMENTO_DOMICILIAR = ""
                    Else
                        vDEPARTAMENTO_DOMICILIAR = MiDataTable1.Rows(X).Item(45).ToString
                    End If
                    If MiDataTable1.Rows(X).Item(19).ToString = "." Or MiDataTable1.Rows(X).Item(19).ToString = "0" Then
                        vTELEFONO = ""
                    End If
                    If MiDataTable1.Rows(X).Item(18).ToString = "." Or MiDataTable1.Rows(X).Item(18).ToString = "0" Or Len(MiDataTable1.Rows(X).Item(18).ToString) <> 8 Then
                        vCELULAR = ""
                    Else
                        vCELULAR = MiDataTable1.Rows(X).Item(18).ToString
                    End If
                    If MiDataTable1.Rows(X).Item(17).ToString = "." Or MiDataTable1.Rows(X).Item(17).ToString = "0" Or Len(MiDataTable1.Rows(X).Item(17).ToString) <> 8 Then
                        vCELULAR = vCELULAR
                    Else
                        If vCELULAR <> "" Then
                            vCELULAR = vCELULAR & ", " & MiDataTable1.Rows(X).Item(17).ToString
                        Else
                            vCELULAR = MiDataTable1.Rows(X).Item(17).ToString
                        End If
                    End If
                    vCORREO_ELECTRONICO = MiDataTable1.Rows(X).Item(20).ToString
                    If Len(vCORREO_ELECTRONICO) < 5 Then
                        vCORREO_ELECTRONICO = ""
                    End If
                    If Len(MiDataTable1.Rows(X).Item(21).ToString) < 5 Then
                        If Len(vCORREO_ELECTRONICO) < 5 Then
                            vCORREO_ELECTRONICO = MiDataTable1.Rows(X).Item(21).ToString
                        Else
                            If Len(MiDataTable1.Rows(X).Item(21).ToString) > 5 Then
                                vCORREO_ELECTRONICO = vCORREO_ELECTRONICO & ", " & MiDataTable1.Rows(X).Item(21).ToString
                            End If
                        End If
                    End If
                    If Len(vCORREO_ELECTRONICO) < 5 Then
                        vCORREO_ELECTRONICO = ""
                    End If
                    If MiDataTable1.Rows(X).Item(13).ToString = "" Or MiDataTable1.Rows(X).Item(13).ToString = "." Or MiDataTable1.Rows(X).Item(13).ToString = "0" Then
                        vTIPO_IDENTIFICACION = "SIN DEFINIR"
                        vID = ""
                    Else
                        If Len(MiDataTable1.Rows(X).Item(13).ToString) > 10 Then
                            vTIPO_IDENTIFICACION = "CEDULA IDENTIDAD"
                            vID = Replace(MiDataTable1.Rows(X).Item(13).ToString, "-", "")
                        Else
                            vTIPO_IDENTIFICACION = "SIN DEFINIR"
                            vID = ""
                        End If
                    End If

                    If Not IsDBNull(MiDataTable1.Rows(X).Item(88).ToString) Then : vNA_BENEFICIARIO = MiDataTable1.Rows(X).Item(88).ToString : Else : vNA_BENEFICIARIO = "" : End If
                    If Not IsDBNull(MiDataTable1.Rows(X).Item(89).ToString) Then : vPAIS_EMISION_ID = MiDataTable1.Rows(X).Item(89).ToString : Else : vPAIS_EMISION_ID = "" : End If
                    If Not IsDBNull(MiDataTable1.Rows(X).Item(90).ToString) Then : vFEC_EMISION_ID = MiDataTable1.Rows(X).Item(90).ToString : Else : vFEC_EMISION_ID = "" : End If
                    If Not IsDBNull(MiDataTable1.Rows(X).Item(91).ToString) Then : vFEC_VENCIMIENTO_ID = MiDataTable1.Rows(X).Item(91).ToString : Else : vFEC_VENCIMIENTO_ID = "" : End If

                    Me.Cursor = Cursors.Default
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
    Dim vNOMBRE_ESTUDIO As String
    Private Sub BUSCAR_OFICIO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable1 As New DataTable
        Dim MiDataAdapter1 As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter1 = New SqlDataAdapter("SELECT * FROM [MAESTRO DE ESTUDIOS] WHERE ID_M_P = " & vID_M_P & " AND ID_T_ESTUDIO = 8 ORDER BY ID_M_P", Conexion.CxRRHH)
            MiDataTable1.Clear()
            MiDataAdapter1.Fill(MiDataTable1)
            If MiDataTable1.Rows.Count > 0 Then
                If Len(MiDataTable1.Rows(0).Item(11).ToString) > 5 Then
                    vNOMBRE_ESTUDIO = MiDataTable1.Rows(0).Item(11).ToString
                Else
                    vNOMBRE_ESTUDIO = ""
                End If
            Else
                vNOMBRE_ESTUDIO = ""
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim C As String = "ACTUALIZADO, PARAMETRO1, PARAMETRO2, PARAMETRO3, PARAMETRO4, PARAMETRO5, PARAMETRO6, PRIMER_NOMBRE, SEGUNDO_NOMBRE, PRIMER_APELLIDO, " &
                      "SEGUNDO_APELLIDO, SEXO, ESTADO_CIVIL, CONYUGE, " &
                      "CELULAR_CONYUGE, CANT_DEPENDIENTES, OCUPACION, PROFESION_OFICIO, FECHA_INGRESO, INGRESO, REFERENCIA, BENEFICIARIO, FECHA_NACIMIENTO," &
                      "PAIS_NACIMIENTO, NACIONALIDAD, PAIS_RESIDENCIA, DIRECCION_DOMICILIAR, MUNICIPIO_DOMICILIAR, DEPARTAMENTO_DOMICILIAR, TELEFONO, CELULAR, " &
                      "CORREO_ELECTRONICO, TIPO_IDENTIFICACION, ID, PAIS_EMISION_ID, FECHA_EMISION_ID, FECHA_VENCE_ID, N1, N2, N3, N4, N5, N6, N7, N8, " &
                      "GRADO, MILITAR, PAME, JEFE, ORDEN_ESTRUCTURA"
    Dim fFECHA_I1, fFECHA_I2, fFECHA_I3, fFECHA_I4 As String
    Private Sub INGRESAR_REGISTRO()

        fFECHA_I1 = Mid(vFECHA_INGRESO, 1, 10)
        Dim f1 As Object = If(fFECHA_I1 <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_I1 + "', 105), 23)", "NULL")

        fFECHA_I2 = Mid(vFECHA_NACIMIENTO, 1, 10)
        Dim f2 As Object = If(fFECHA_I2 <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_I2 + "', 105), 23)", "NULL")

        fFECHA_I3 = Mid(vFEC_EMISION_ID, 1, 10)
        Dim f3 As Object = If(fFECHA_I3 <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_I3 + "', 105), 23)", "NULL")

        fFECHA_I4 = Mid(vFEC_VENCIMIENTO_ID, 1, 10)
        Dim f4 As Object = If(fFECHA_I4 <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFECHA_I4 + "', 105), 23)", "NULL")

        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO BAC] (" & C & ")" &
                                  "values ('FALSE', NULL, NULL, NULL, NULL, NULL, NULL, '" & vN1 & "','" & vN2 & "','" & vA1 & "', " &
                                  "'" & vA2 & "', '" & vSEXO & "', '" & vESTADO_CIVIL & "', '" & vCONYUGE & "', " &
                                  "'" & vCELULAR_CONYUGE & "', " & vCANT_DEPENDIENTES & ", '" & vOCUPACION & "', '" & vPROFESION_OFICIO & "'," & f1 & ", '" & vINGRESO & "', '" & vREFERENCIA & "', '" & vNA_BENEFICIARIO & "', " & f2 & ", " &
                                  "'" & vPAIS_NACIMIENTO & "', '" & vNACIONALIDAD & "', '" & vPAIS_RESIDENCIA & "', '" & vDIRECCION_DOMICILIAR & "', '" & vMUNICIPIO_DOMICILIAR & "', '" & vDEPARTAMENTO_DOMICILIAR & "', '" & vTELEFONO & "', '" & vCELULAR & "', " &
                                  "'" & vCORREO_ELECTRONICO & "', '" & vTIPO_IDENTIFICACION & "', '" & vID & "', '" & vPAIS_EMISION_ID & "', " & f3 & ", " & f4 & ", '" & bacN1 & "', '" & bacN2 & "', '" & bacN3 & "', '" & bacN4 & "', '" & bacN5 & "', '" & bacN6 & "', '" & bacN7 & "', '" & bacN8 & "', " &
                                  "'" & vGRADO & "', '" & vMILITAR & "', '" & vPAME & "', '" & vJEFE & "', " & vORDEN_ESTRUCTURA & ")"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim P1, P2, P3, P4, P5, P6 As String
    Dim ArrCadena As String()
    Dim POSa, POSn As Integer
    Private Sub DIVIDIR_NA()
        Dim LEN_A As Integer
        LEN_A = Len(vA)
        POSa = InStr(vA, " ")
        If POSa = 0 Then
            vA1 = Trim(vA)
        Else
            vA1 = Trim(Mid(vA, 1, POSa))
        End If
        If POSa = 0 Then
            vA2 = ""
        Else
            vA2 = Trim(Mid(vA, POSa, LEN_A))
        End If


        Dim LEN_N As Integer
        LEN_N = Len(vN)
        POSn = InStr(vN, " ")
        If POSn = 0 Then
            vN1 = Trim(vN)
        Else
            vN1 = Trim(Mid(vN, 1, POSn))
        End If
        If POSn = 0 Then
            vN2 = ""
        Else
            vN2 = Trim(Mid(vN, POSn, Len(vN)))
        End If
    End Sub
End Class