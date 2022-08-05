Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Estructura_Org_Importar
    Private Sub Frm_Estructura_Org_Importar_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            NIVEL1x = Me.ComboBox1.Text
            NIVEL2x = Me.ComboBox2.Text
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Estructura_Org_Importar_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RemoveHandler Me.ComboBox2.SelectedIndexChanged, AddressOf ComboBox2_SelectedIndexChanged
        RemoveHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
        Call CARGAR_COMBO_SPYC_CAT_DE_ESTRUCTURA_NIVEL_1()
        If NIVEL1x <> "" Then
            Me.ComboBox1.Text = NIVEL1x
        End If
        Call CARGAR_COMBO_SPYC_CAT_DE_ESTRUCTURA_NIVEL_2()
        If NIVEL2x <> "" Then
            Me.ComboBox2.Text = NIVEL2x
        End If
        CADENAsql1 = "Select * from [SPYC_VISTA MAESTRO DE CARGOS] WHERE ID_N1 = " & Me.ComboBox1.SelectedValue & " AND ID_N2 = " & Me.ComboBox2.SelectedValue & " ORDER BY ID_N1, ID_N2, ORDEN"
        Call CARGAR_DATAGRIDVIEW_1()
        Call ARMAR_DATAGRIDVIEW_1()
        AddHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
        AddHandler Me.ComboBox2.SelectedIndexChanged, AddressOf ComboBox2_SelectedIndexChanged
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        NIVEL1x = Me.ComboBox1.Text
        NIVEL2x = Me.ComboBox2.Text
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_SPYC_CAT_DE_ESTRUCTURA_NIVEL_1()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [SPYC_CAT DE ESTRUCTURA NIVEL 1] order by ORDEN", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
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
    Private Sub CARGAR_COMBO_SPYC_CAT_DE_ESTRUCTURA_NIVEL_2()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [SPYC_CAT DE ESTRUCTURA NIVEL 2] WHERE ID_N1 = " & Me.ComboBox1.SelectedValue & " order by ORDEN", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox2
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_N2"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_1()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql1, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[SPYC_VISTA MAESTRO DE CARGOS]")
        DGV3.DataSource = MiDataSet.Tables("[SPYC_VISTA MAESTRO DE CARGOS]")
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_1()
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
        Me.DGV3.Columns(7).Width = 110
        Me.DGV3.Columns(7).Visible = True
        Me.DGV3.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(8).HeaderText = "ID_T_ES"
        Me.DGV3.Columns(8).Width = 0
        Me.DGV3.Columns(8).Visible = False

        Me.DGV3.Columns(9).HeaderText = "D_T_ES"
        Me.DGV3.Columns(9).Width = 0
        Me.DGV3.Columns(9).Visible = False

        Me.DGV3.Columns(10).HeaderText = "No. Orden [MIXTO]"
        Me.DGV3.Columns(10).Width = 50
        Me.DGV3.Columns(10).Visible = True
        Me.DGV3.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV3.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV3.Columns(10).DefaultCellStyle.BackColor = Color.Khaki

        Me.DGV3.Columns(11).HeaderText = "No. Orden [MILITAR]"
        Me.DGV3.Columns(11).Width = 50
        Me.DGV3.Columns(11).Visible = True
        Me.DGV3.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV3.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV3.Columns(11).DefaultCellStyle.BackColor = Color.LightGreen

        Me.DGV3.Columns(12).HeaderText = "No. Orden [PAME]"
        Me.DGV3.Columns(12).Width = 50
        Me.DGV3.Columns(12).Visible = True
        Me.DGV3.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV3.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV3.Columns(12).DefaultCellStyle.BackColor = Color.LightBlue

        Me.DGV3.Columns(13).HeaderText = "ID_CARGO_ES"
        Me.DGV3.Columns(13).Width = 0
        Me.DGV3.Columns(13).Visible = False

        Me.DGV3.Columns(14).HeaderText = "Cargo"
        Me.DGV3.Columns(14).Width = 155
        Me.DGV3.Columns(14).Visible = True
        Me.DGV3.Columns(14).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(15).HeaderText = "ID_GP"
        Me.DGV3.Columns(15).Width = 0
        Me.DGV3.Columns(15).Visible = False

        Me.DGV3.Columns(16).HeaderText = "GP"
        Me.DGV3.Columns(16).Width = 80
        Me.DGV3.Columns(16).Visible = True
        Me.DGV3.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(17).HeaderText = "ID_GR"
        Me.DGV3.Columns(17).Width = 0
        Me.DGV3.Columns(17).Visible = False

        Me.DGV3.Columns(18).HeaderText = "GR"
        Me.DGV3.Columns(18).Width = 80
        Me.DGV3.Columns(18).Visible = True
        Me.DGV3.Columns(18).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(19).HeaderText = "ID_CAT_SALARIAL"
        Me.DGV3.Columns(19).Width = 0
        Me.DGV3.Columns(19).Visible = False

        Me.DGV3.Columns(20).HeaderText = "D_CAT_SALARIAL"
        Me.DGV3.Columns(20).Width = 0
        Me.DGV3.Columns(20).Visible = False

        Me.DGV3.Columns(21).HeaderText = "ID_M_P"
        Me.DGV3.Columns(21).Width = 0
        Me.DGV3.Columns(21).Visible = False

        Me.DGV3.Columns(22).HeaderText = "Codigo"
        Me.DGV3.Columns(22).Width = 70
        Me.DGV3.Columns(22).Visible = True
        Me.DGV3.Columns(22).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(23).HeaderText = "Apellidos y Empleados"
        Me.DGV3.Columns(23).Width = 175
        Me.DGV3.Columns(23).Visible = True
        Me.DGV3.Columns(23).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV3.Columns(23).DefaultCellStyle.BackColor = Color.Khaki

        Me.DGV3.Columns(24).HeaderText = "ID_SITUACION"
        Me.DGV3.Columns(24).Width = 0
        Me.DGV3.Columns(24).Visible = False

        Me.DGV3.Columns(25).HeaderText = "Jefe"
        Me.DGV3.Columns(25).Width = 0
        Me.DGV3.Columns(25).Visible = False
        Me.DGV3.Columns(25).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(26).HeaderText = "Mixta"
        Me.DGV3.Columns(26).Width = 0
        Me.DGV3.Columns(26).Visible = False
        Me.DGV3.Columns(26).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(27).HeaderText = "Militar"
        Me.DGV3.Columns(27).Width = 40
        Me.DGV3.Columns(27).Visible = True
        Me.DGV3.Columns(27).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV3.Columns(28).HeaderText = "Pame"
        Me.DGV3.Columns(28).Width = 40
        Me.DGV3.Columns(28).Visible = True
        Me.DGV3.Columns(28).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Call CARGAR_COMBO_SPYC_CAT_DE_ESTRUCTURA_NIVEL_2()
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        RemoveHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
        CADENAsql1 = "Select * from [SPYC_VISTA MAESTRO DE CARGOS] WHERE ID_N1 = " & Me.ComboBox1.SelectedValue & " AND ID_N2 = " & Me.ComboBox2.SelectedValue & " ORDER BY ID_N1, ID_N2, ORDEN"
        Call CARGAR_DATAGRIDVIEW_1()
        Call ARMAR_DATAGRIDVIEW_1()
        AddHandler Me.ComboBox1.SelectedIndexChanged, AddressOf ComboBox1_SelectedIndexChanged
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.DGV3.RowCount = 0 Then
            MsgBox("No existen registros seleccionados", vbInformation, "Mensaje del Sistema")
            Me.DGV3.Focus()
            Exit Sub
        End If
        Call RECORRER_DGV()
        Call CARGAR_DATAGRIDVIEW_3()
        Call ARMAR_DATAGRIDVIEW_3()
        NIVEL1x = Me.ComboBox1.Text
        NIVEL2x = Me.ComboBox2.Text
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub RECORRER_DGV()
        If Me.DGV3.RowCount <> 0 Then
            For I = 0 To Me.DGV3.RowCount - 1
                If DGV3.Rows(I).Selected = True Then
                    iID_M_C = Me.DGV3.Item(0, I).Value
                    iID_M_P = Me.DGV3.Item(21, I).Value
                    Call BUSCAR_DATO_DE_PERSONA()
                    If EXISTE_PERSONA = True Then
                        Call BUSCAR_DATO_DE_PERSONA_EN_LOCAL()
                        If EXISTE_PERSONA_LOCAL = False Then
                            Call AGREGAR_REGISTRO_MAESTRO_DE_PERSONAS()
                        End If
                        Call BUSCAR_DATO_DE_CARGO()
                        Call AGREGAR_REGISTRO_MAESTRO_DE_CARGO()
                    End If
                End If
            Next
        End If
    End Sub
    Dim cadena_CAMPOS1, cadena_VALORES1 As String
    Private Sub AGREGAR_REGISTRO_MAESTRO_DE_PERSONAS()
        cadena_CAMPOS1 = "ID_M_P, FEC_NAC, FEC_ING_PAME, FEC_ING_EN, " &
                        "CODIGO, APELLIDOS, NOMBRES, ID_T_SEXO, DIRECCION_1, " &
                        "DIRECCION_2, NSS, N_CEDULA, N_LIC, N_MINSA, " &
                        "LIC_IMAGEN, TELEFONO_1, TELEFONO_2, CONVENCIONAL, CORREO_1, " &
                        "CORREO_2, WHATSAPP, FACEBOOK, TWITTER, PESO, " &
                        "ESTATURA, ID_TEZ, ID_CABELLO, ID_N_ACADEMICO, ID_N_PROFESIONAL, " &
                        "ID_NACIONALIDAD_N, ID_DPTO_N, ID_M_N, LUGAR_N, ID_NACIONALIDAD_D, " &
                        "ID_DPTO_D, ID_M_D, LUGAR_D, ID_E_CIVIL, ID_T_HORARIO, " &
                        "N_C_BANCO, EMPLEADO_ANT, ID_T_AFILIACION, ID_T_PLAZA, D_N1, " &
                        "D_N2, FOTO, ID_ESTADO_P, BECADO, ID_C_LIC, " &
                        "ID_T_SANGUINEO, OBSERVACIONES, ID_IPSS, ESPECIALISTA, SUBESPECIALISTA, " &
                        "ID_F, EN_SITUACION, CARGO_REAL, JEFE_REAL, CALCULO_VACACIONAL, " &
                        "ID_N_COMPETENCIA, SERVICIOS_PROFESIONALES, ID_ADMIN"

        Dim ms As New System.IO.MemoryStream() 'creando el memoryStream
        Me.PictureBox1.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)

        Dim fFEC_ING_EN = Mid(i_mp_FEC_ING_EN, 1, 10)
        Dim fFEC_NAC As String = Mid(i_mp_FEC_NAC, 1, 10)
        Dim fFEC_ING As String = Mid(i_mp_FEC_ING_PAME, 1, 10)

        Dim fechaNacimiento = "Convert(VARCHAR(10), Convert(Date, '" + fFEC_NAC + "', 105), 23)"
        Dim fechaingreso = "Convert(VARCHAR(10), Convert(Date, '" + fFEC_ING + "', 105), 23)"
        Dim fechaingresoEN As Object = If(fFEC_ING_EN <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFEC_ING_EN + "', 105), 23)", "NULL")


        cadena_VALORES1 = "" & i_mp_ID_M_P & ", " & fechaNacimiento & ", " & fechaingreso & ", " & fechaingresoEN & ", " &
                         "'" & i_mp_CODIGO & "', '" & i_mp_APELLIDOS & "', '" & i_mp_NOMBRES & "', " & i_mp_ID_T_SEXO & ", '" & i_mp_DIRECCION_1 & "', " &
                         "'" & i_mp_DIRECCION_2 & "', '" & i_mp_NSS & "', '" & i_mp_N_CEDULA & "', '" & i_mp_N_LIC & "', '" & i_mp_N_MINSA & "', " &
                         "'" & i_mp_LIC_IMAGEN & "', '" & i_mp_TELEFONO_1 & "', '" & i_mp_TELEFONO_2 & "', '" & i_mp_CONVENCIONAL & "', '" & i_mp_CORREO_1 & "', " &
                         "'" & i_mp_CORREO_2 & "', '" & i_mp_WHATSAPP & "', '" & i_mp_FACEBOOK & "', '" & i_mp_TWITTER & "', '" & i_mp_PESO & "', " &
                         "'" & i_mp_ESTATURA & "', " & i_mp_ID_TEZ & ", " & i_mp_ID_CABELLO & ", " & i_mp_ID_N_ACADEMICO & ", " & i_mp_ID_N_PROFESIONAL & ", " &
                         "" & i_mp_ID_NACIONALIDAD_N & ", " & i_mp_ID_DPTO_N & ", " & i_mp_ID_M_N & ", '" & i_mp_LUGAR_N & "', " & i_mp_ID_NACIONALIDAD_D & ", " &
                         "" & i_mp_ID_DPTO_D & ", " & i_mp_ID_M_D & ", '" & i_mp_LUGAR_D & "', " & i_mp_ID_E_CIVIL & ", " & i_mp_ID_T_HORARIO & ", " &
                         "'" & i_mp_N_C_BANCO & "', '" & i_mp_EMPLEADO_ANT & "', " & i_mp_ID_T_AFILIACION & ", " & i_mp_ID_T_PLAZA & ", '" & i_mp_D_N1 & "', " &
                         "'" & i_mp_D_N2 & "',  @data, '" & i_mp_ID_ESTADO_P & "', '" & i_mp_BECADO & "', '" & i_mp_ID_C_LIC & "', " &
                         "" & i_mp_ID_T_SANGUINEO & ", '" & i_mp_OBSERVACIONES & "', " & i_mp_ID_IPSS & ", '" & i_mp_ESPECIALISTA & "', '" & i_mp_SUBESPECIALISTA & "', " &
                         "" & i_mp_ID_F & ", '" & i_mp_EN_SITUACION & "', '" & i_mp_CARGO_REAL & "', '" & i_mp_JEFE_REAL & "', '" & i_mp_CALCULO_VACACIONAL & "', " &
                         "" & i_mp_ID_N_COMPETENCIA & ", '" & i_mp_SERVICIOS_PROFESIONALES & "', 1"

        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim SQL As String
        Try
            SQL = "INSERT INTO [MAESTRO DE PERSONAS] (" & cadena_CAMPOS1 & ") values (" & cadena_VALORES1 & ")"
            Dim cmd As New SqlCommand(SQL, CxRRHH)
            cmd.Parameters.Add(New SqlParameter("@data", MS.GetBuffer()))
            cmd.ExecuteNonQuery()
            cmd = Nothing
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim EXISTE_PERSONA_LOCAL As Boolean
    Private Sub BUSCAR_DATO_DE_PERSONA_EN_LOCAL()
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [MAESTRO DE PERSONAS] WHERE ID_M_P = " & iID_M_P & " AND ID_ESTADO_P = 1 ORDER BY ID_M_P", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                EXISTE_PERSONA_LOCAL = True
            Else
                EXISTE_PERSONA_LOCAL = False
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim EXISTE_PERSONA As Boolean
    Private Sub BUSCAR_DATO_DE_PERSONA()
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [SPYC_MAESTRO DE PERSONAS] WHERE ID_M_P = " & iID_M_P & " AND ID_ESTADO_P = 1 ORDER BY ID_M_P", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                i_mp_ID_M_P = MiDataTable.Rows(0).Item(0).ToString
                i_mp_FEC_NAC = MiDataTable.Rows(0).Item(1).ToString
                i_mp_FEC_ING_PAME = MiDataTable.Rows(0).Item(2).ToString
                i_mp_FEC_ING_EN = MiDataTable.Rows(0).Item(3).ToString
                i_mp_CODIGO = MiDataTable.Rows(0).Item(4).ToString
                i_mp_APELLIDOS = MiDataTable.Rows(0).Item(5).ToString
                i_mp_NOMBRES = MiDataTable.Rows(0).Item(6).ToString
                i_mp_ID_T_SEXO = MiDataTable.Rows(0).Item(7).ToString
                i_mp_DIRECCION_1 = MiDataTable.Rows(0).Item(8).ToString
                i_mp_DIRECCION_2 = MiDataTable.Rows(0).Item(9).ToString
                i_mp_NSS = MiDataTable.Rows(0).Item(10).ToString
                i_mp_N_CEDULA = MiDataTable.Rows(0).Item(11).ToString
                i_mp_N_LIC = MiDataTable.Rows(0).Item(12).ToString
                i_mp_N_MINSA = MiDataTable.Rows(0).Item(13).ToString
                i_mp_LIC_IMAGEN = MiDataTable.Rows(0).Item(14).ToString
                i_mp_TELEFONO_1 = MiDataTable.Rows(0).Item(15).ToString
                i_mp_TELEFONO_2 = MiDataTable.Rows(0).Item(16).ToString
                i_mp_CONVENCIONAL = MiDataTable.Rows(0).Item(17).ToString
                i_mp_CORREO_1 = MiDataTable.Rows(0).Item(18).ToString
                i_mp_CORREO_2 = MiDataTable.Rows(0).Item(19).ToString
                i_mp_WHATSAPP = MiDataTable.Rows(0).Item(20).ToString
                i_mp_FACEBOOK = MiDataTable.Rows(0).Item(21).ToString
                i_mp_TWITTER = MiDataTable.Rows(0).Item(22).ToString
                i_mp_PESO = MiDataTable.Rows(0).Item(23).ToString
                i_mp_ESTATURA = MiDataTable.Rows(0).Item(24).ToString
                i_mp_ID_TEZ = MiDataTable.Rows(0).Item(25).ToString
                i_mp_ID_CABELLO = MiDataTable.Rows(0).Item(26).ToString
                i_mp_ID_N_ACADEMICO = MiDataTable.Rows(0).Item(27).ToString
                i_mp_ID_N_PROFESIONAL = MiDataTable.Rows(0).Item(28).ToString
                i_mp_ID_NACIONALIDAD_N = MiDataTable.Rows(0).Item(29).ToString
                i_mp_ID_DPTO_N = MiDataTable.Rows(0).Item(30).ToString
                i_mp_ID_M_N = MiDataTable.Rows(0).Item(31).ToString
                i_mp_LUGAR_N = MiDataTable.Rows(0).Item(32).ToString
                i_mp_ID_NACIONALIDAD_D = MiDataTable.Rows(0).Item(33).ToString
                i_mp_ID_DPTO_D = MiDataTable.Rows(0).Item(34).ToString
                i_mp_ID_M_D = MiDataTable.Rows(0).Item(35).ToString
                i_mp_LUGAR_D = MiDataTable.Rows(0).Item(36).ToString
                i_mp_ID_E_CIVIL = MiDataTable.Rows(0).Item(37).ToString
                i_mp_ID_T_HORARIO = MiDataTable.Rows(0).Item(38).ToString
                i_mp_N_C_BANCO = MiDataTable.Rows(0).Item(39).ToString
                i_mp_EMPLEADO_ANT = MiDataTable.Rows(0).Item(40).ToString
                i_mp_ID_T_AFILIACION = MiDataTable.Rows(0).Item(41).ToString
                i_mp_ID_T_PLAZA = MiDataTable.Rows(0).Item(42).ToString
                i_mp_D_N1 = MiDataTable.Rows(0).Item(43).ToString
                i_mp_D_N2 = MiDataTable.Rows(0).Item(44).ToString
                Call CARGAR_FOTO()
                i_mp_ID_ESTADO_P = MiDataTable.Rows(0).Item(46).ToString
                i_mp_BECADO = MiDataTable.Rows(0).Item(47).ToString
                i_mp_ID_C_LIC = MiDataTable.Rows(0).Item(48).ToString
                i_mp_ID_T_SANGUINEO = MiDataTable.Rows(0).Item(49).ToString
                i_mp_OBSERVACIONES = MiDataTable.Rows(0).Item(50).ToString
                i_mp_ID_IPSS = MiDataTable.Rows(0).Item(51).ToString
                i_mp_ESPECIALISTA = MiDataTable.Rows(0).Item(52).ToString
                i_mp_SUBESPECIALISTA = MiDataTable.Rows(0).Item(53).ToString
                i_mp_ID_F = MiDataTable.Rows(0).Item(54).ToString
                i_mp_EN_SITUACION = MiDataTable.Rows(0).Item(55).ToString
                i_mp_CARGO_REAL = MiDataTable.Rows(0).Item(56).ToString
                i_mp_JEFE_REAL = MiDataTable.Rows(0).Item(57).ToString
                i_mp_CALCULO_VACACIONAL = MiDataTable.Rows(0).Item(58).ToString
                i_mp_ID_N_COMPETENCIA = MiDataTable.Rows(0).Item(59).ToString
                i_mp_SERVICIOS_PROFESIONALES = MiDataTable.Rows(0).Item(60).ToString
                EXISTE_PERSONA = True
            Else
                EXISTE_PERSONA = False
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_FOTO()
        Dim DT As New DataTable
        Dim DA As New SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            CADENAsql = "SELECT * FROM [SPYC_MAESTRO DE PERSONAS] WHERE ID_M_P = " & iID_M_P & " ORDER BY ID_M_P"
            DA = New SqlDataAdapter(CADENAsql, CxRRHH)
            DA.Fill(DT)
            If DT.Rows.Count > 0 Then
                Dim ms As New System.IO.MemoryStream()
                Dim imageBuffer() As Byte = CType(DT.Rows(0).Item("FOTO"), Byte())
                ms = New System.IO.MemoryStream(imageBuffer)
                Me.PictureBox1.Image = Nothing
                Me.PictureBox1.Image = (Image.FromStream(ms))
                Me.PictureBox1.BackgroundImageLayout = ImageLayout.Stretch
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
        DT.Clear()
        DT.Reset()
    End Sub
    Private Sub BUSCAR_DATO_DE_CARGO()
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [SPYC_MAESTRO DE CARGOS] WHERE ID_M_C = " & iID_M_C & " ORDER BY ID_N1, ID_N2, ORDEN", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                i_mc_ID_M_C = MiDataTable.Rows(0).Item(0).ToString
                i_mc_ORDEN = MiDataTable.Rows(0).Item(1).ToString
                i_mc_ID_N1 = MiDataTable.Rows(0).Item(2).ToString
                i_mc_ID_N2 = MiDataTable.Rows(0).Item(3).ToString
                i_mc_ID_T_ES = MiDataTable.Rows(0).Item(4).ToString
                i_mc_N_ORDEN_MIXTA = MiDataTable.Rows(0).Item(5).ToString
                i_mc_N_ORDEN_MILITAR = MiDataTable.Rows(0).Item(6).ToString
                i_mc_N_ORDEN_PAME = MiDataTable.Rows(0).Item(7).ToString
                i_mc_ID_CARGO_ES = MiDataTable.Rows(0).Item(8).ToString
                i_mc_ID_GP = MiDataTable.Rows(0).Item(9).ToString
                i_mc_ID_GR = MiDataTable.Rows(0).Item(10).ToString
                i_mc_ID_CAT_SALARIAL = MiDataTable.Rows(0).Item(11).ToString
                i_mc_ID_M_P = MiDataTable.Rows(0).Item(12).ToString
                i_mc_ID_SITUACION = MiDataTable.Rows(0).Item(13).ToString
                i_mc_JEFE = MiDataTable.Rows(0).Item(14).ToString
                i_mc_MIXTA = MiDataTable.Rows(0).Item(15).ToString
                i_mc_MILITAR = MiDataTable.Rows(0).Item(16).ToString
                i_mc_PAME = MiDataTable.Rows(0).Item(17).ToString
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim cadena_CAMPOS2, cadena_VALORES2 As String
    Dim vN1, vN2, vN3, vN4, vN5, vN6, vN7, vN8 As String
    Private Sub AGREGAR_REGISTRO_MAESTRO_DE_CARGO()
        cadena_CAMPOS2 = "ID_M_C, ORDEN, ID_N1, ID_N2, ID_N3, " &
                        "ID_N4, ID_N5, ID_N6, ID_N7, ID_N8, ID_T_ES, N_ORDEN_MIXTA, N_ORDEN_MILITAR, " &
                        "N_ORDEN_PAME, ID_CARGO_ES, ID_GP, ID_GR, ID_CAT_SALARIAL, " &
                        "ID_M_P, ID_SITUACION, JEFE, MIXTA, MILITAR, " &
                        "PAME, ID_ESTABLECIMIENTO"
        'NIVEL 1
        If Frm_Estructura_Org.ComboBox1.SelectedValue = 0 Then
            i_mc_ID_N1 = Nothing
        Else
            i_mc_ID_N1 = Frm_Estructura_Org.ComboBox1.SelectedValue
        End If
        'NIVEL 2
        If Frm_Estructura_Org.ComboBox2.SelectedValue = 0 Then
            i_mc_ID_N2 = Nothing
        Else
            i_mc_ID_N2 = Frm_Estructura_Org.ComboBox2.SelectedValue
        End If
        'NIVEL 3
        If Frm_Estructura_Org.ComboBox3.SelectedValue = 0 Then
            i_mc_ID_N3 = Nothing
        Else
            i_mc_ID_N3 = Frm_Estructura_Org.ComboBox3.SelectedValue
        End If
        'NIVEL 4
        If Frm_Estructura_Org.ComboBox4.SelectedValue = 0 Then
            i_mc_ID_N4 = Nothing
        Else
            i_mc_ID_N4 = Frm_Estructura_Org.ComboBox4.SelectedValue
        End If
        'NIVEL 5
        If Frm_Estructura_Org.ComboBox5.SelectedValue = 0 Then
            i_mc_ID_N5 = Nothing
        Else
            i_mc_ID_N5 = Frm_Estructura_Org.ComboBox5.SelectedValue
        End If
        'NIVEL 6
        If Frm_Estructura_Org.ComboBox6.SelectedValue = 0 Then
            i_mc_ID_N6 = Nothing
        Else
            i_mc_ID_N6 = Frm_Estructura_Org.ComboBox6.SelectedValue
        End If
        'NIVEL 7
        If Frm_Estructura_Org.ComboBox7.SelectedValue = 0 Then
            i_mc_ID_N7 = Nothing
        Else
            i_mc_ID_N7 = Frm_Estructura_Org.ComboBox7.SelectedValue
        End If
        'NIVEL 8
        If Frm_Estructura_Org.ComboBox8.SelectedValue = 0 Then
            i_mc_ID_N8 = Nothing
        Else
            i_mc_ID_N8 = Frm_Estructura_Org.ComboBox8.SelectedValue
        End If

        If i_mc_ID_N1 = 0 Then : vN1 = "NULL" : Else : vN1 = Frm_Estructura_Org.ComboBox1.SelectedValue : End If
        If i_mc_ID_N2 = 0 Then : vN2 = "NULL" : Else : vN2 = Frm_Estructura_Org.ComboBox2.SelectedValue : End If
        If i_mc_ID_N3 = 0 Then : vN3 = "NULL" : Else : vN3 = Frm_Estructura_Org.ComboBox3.SelectedValue : End If
        If i_mc_ID_N4 = 0 Then : vN4 = "NULL" : Else : vN4 = Frm_Estructura_Org.ComboBox4.SelectedValue : End If
        If i_mc_ID_N5 = 0 Then : vN5 = "NULL" : Else : vN5 = Frm_Estructura_Org.ComboBox5.SelectedValue : End If
        If i_mc_ID_N6 = 0 Then : vN6 = "NULL" : Else : vN6 = Frm_Estructura_Org.ComboBox6.SelectedValue : End If
        If i_mc_ID_N7 = 0 Then : vN7 = "NULL" : Else : vN7 = Frm_Estructura_Org.ComboBox7.SelectedValue : End If
        If i_mc_ID_N8 = 0 Then : vN8 = "NULL" : Else : vN8 = Frm_Estructura_Org.ComboBox8.SelectedValue : End If

        cadena_VALORES2 = "" & i_mc_ID_M_C & ", '" & i_mc_ORDEN & "', " & vN1 & ", " & vN2 & ", " & vN3 & ", " &
                          "" & vN4 & ", " & vN5 & ", " & vN6 & ", " & vN7 & ", " & vN8 & ", " & i_mc_ID_T_ES & ", '" & i_mc_N_ORDEN_MIXTA & "', '" & i_mc_N_ORDEN_MILITAR & "', " &
                          "'" & i_mc_N_ORDEN_PAME & "', " & i_mc_ID_CARGO_ES & ", " & i_mc_ID_GP & ", " & i_mc_ID_GR & ", " & i_mc_ID_CAT_SALARIAL & ", " &
                          "" & i_mc_ID_M_P & ", " & i_mc_ID_SITUACION & ", '" & i_mc_JEFE & "', '" & i_mc_MIXTA & "', '" & i_mc_MILITAR & "', " &
                          "'" & i_mc_PAME & "', 1"

        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "INSERT INTO [MAESTRO DE CARGOS] (" & cadena_CAMPOS2 & ")" &
                                                                "values (" & cadena_VALORES2 & ")"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW_3()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql3, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE CARGOS SIT]")
        Frm_Estructura_Org.DGV3.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE CARGOS SIT]")
        Frm_Estructura_Org.Label4.Text = Frm_Estructura_Org.DGV3.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW_3()
        Frm_Estructura_Org.DGV3.Columns(0).HeaderText = "Id"
        Frm_Estructura_Org.DGV3.Columns(0).Width = 20
        Frm_Estructura_Org.DGV3.Columns(0).Visible = True
        Frm_Estructura_Org.DGV3.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Estructura_Org.DGV3.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Frm_Estructura_Org.DGV3.Columns(1).HeaderText = "ORDEN"
        Frm_Estructura_Org.DGV3.Columns(1).Width = 0
        Frm_Estructura_Org.DGV3.Columns(1).Visible = False

        Frm_Estructura_Org.DGV3.Columns(2).HeaderText = "ID_N1"
        Frm_Estructura_Org.DGV3.Columns(2).Width = 0
        Frm_Estructura_Org.DGV3.Columns(2).Visible = False

        Frm_Estructura_Org.DGV3.Columns(3).HeaderText = "ORDEN_N1"
        Frm_Estructura_Org.DGV3.Columns(3).Width = 0
        Frm_Estructura_Org.DGV3.Columns(3).Visible = False

        Frm_Estructura_Org.DGV3.Columns(4).HeaderText = "Nivel 1"
        Frm_Estructura_Org.DGV3.Columns(4).Width = 0
        Frm_Estructura_Org.DGV3.Columns(4).Visible = False
        Frm_Estructura_Org.DGV3.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(5).HeaderText = "ID_N2"
        Frm_Estructura_Org.DGV3.Columns(5).Width = 0
        Frm_Estructura_Org.DGV3.Columns(5).Visible = False

        Frm_Estructura_Org.DGV3.Columns(6).HeaderText = "ORDEN_N2"
        Frm_Estructura_Org.DGV3.Columns(6).Width = 0
        Frm_Estructura_Org.DGV3.Columns(6).Visible = False

        Frm_Estructura_Org.DGV3.Columns(7).HeaderText = "Nivel 2"
        Frm_Estructura_Org.DGV3.Columns(7).Width = 0
        Frm_Estructura_Org.DGV3.Columns(7).Visible = False
        Frm_Estructura_Org.DGV3.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(8).HeaderText = "ID_N3"
        Frm_Estructura_Org.DGV3.Columns(8).Width = 0
        Frm_Estructura_Org.DGV3.Columns(8).Visible = False

        Frm_Estructura_Org.DGV3.Columns(9).HeaderText = "ORDEN_N3"
        Frm_Estructura_Org.DGV3.Columns(9).Width = 0
        Frm_Estructura_Org.DGV3.Columns(9).Visible = False

        Frm_Estructura_Org.DGV3.Columns(10).HeaderText = "Nivel 3"
        Frm_Estructura_Org.DGV3.Columns(10).Width = 0
        Frm_Estructura_Org.DGV3.Columns(10).Visible = False
        Frm_Estructura_Org.DGV3.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(11).HeaderText = "ID_N4"
        Frm_Estructura_Org.DGV3.Columns(11).Width = 0
        Frm_Estructura_Org.DGV3.Columns(11).Visible = False

        Frm_Estructura_Org.DGV3.Columns(12).HeaderText = "ORDEN_N4"
        Frm_Estructura_Org.DGV3.Columns(12).Width = 0
        Frm_Estructura_Org.DGV3.Columns(12).Visible = False

        Frm_Estructura_Org.DGV3.Columns(13).HeaderText = "Nivel 4"
        Frm_Estructura_Org.DGV3.Columns(13).Width = 0
        Frm_Estructura_Org.DGV3.Columns(13).Visible = False
        Frm_Estructura_Org.DGV3.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(14).HeaderText = "ID_N5"
        Frm_Estructura_Org.DGV3.Columns(14).Width = 0
        Frm_Estructura_Org.DGV3.Columns(14).Visible = False

        Frm_Estructura_Org.DGV3.Columns(15).HeaderText = "ORDEN_N5"
        Frm_Estructura_Org.DGV3.Columns(15).Width = 0
        Frm_Estructura_Org.DGV3.Columns(15).Visible = False

        Frm_Estructura_Org.DGV3.Columns(16).HeaderText = "Nivel 5"
        Frm_Estructura_Org.DGV3.Columns(16).Width = 0
        Frm_Estructura_Org.DGV3.Columns(16).Visible = False
        Frm_Estructura_Org.DGV3.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(17).HeaderText = "ID_N6"
        Frm_Estructura_Org.DGV3.Columns(17).Width = 0
        Frm_Estructura_Org.DGV3.Columns(17).Visible = False

        Frm_Estructura_Org.DGV3.Columns(18).HeaderText = "ORDEN_N6"
        Frm_Estructura_Org.DGV3.Columns(18).Width = 0
        Frm_Estructura_Org.DGV3.Columns(18).Visible = False

        Frm_Estructura_Org.DGV3.Columns(19).HeaderText = "Nivel 6"
        Frm_Estructura_Org.DGV3.Columns(19).Width = 0
        Frm_Estructura_Org.DGV3.Columns(19).Visible = False
        Frm_Estructura_Org.DGV3.Columns(19).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(20).HeaderText = "ID_N7"
        Frm_Estructura_Org.DGV3.Columns(20).Width = 0
        Frm_Estructura_Org.DGV3.Columns(20).Visible = False

        Frm_Estructura_Org.DGV3.Columns(21).HeaderText = "ORDEN_N7"
        Frm_Estructura_Org.DGV3.Columns(21).Width = 0
        Frm_Estructura_Org.DGV3.Columns(21).Visible = False

        Frm_Estructura_Org.DGV3.Columns(22).HeaderText = "Nivel 7"
        Frm_Estructura_Org.DGV3.Columns(22).Width = 0
        Frm_Estructura_Org.DGV3.Columns(22).Visible = False
        Frm_Estructura_Org.DGV3.Columns(22).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(23).HeaderText = "ID_N8"
        Frm_Estructura_Org.DGV3.Columns(23).Width = 0
        Frm_Estructura_Org.DGV3.Columns(23).Visible = False

        Frm_Estructura_Org.DGV3.Columns(24).HeaderText = "ORDEN_N8"
        Frm_Estructura_Org.DGV3.Columns(24).Width = 0
        Frm_Estructura_Org.DGV3.Columns(24).Visible = False

        Frm_Estructura_Org.DGV3.Columns(25).HeaderText = "Nivel 8"
        Frm_Estructura_Org.DGV3.Columns(25).Width = 0
        Frm_Estructura_Org.DGV3.Columns(25).Visible = False
        Frm_Estructura_Org.DGV3.Columns(25).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(26).HeaderText = "ID_T_ES"
        Frm_Estructura_Org.DGV3.Columns(26).Width = 0
        Frm_Estructura_Org.DGV3.Columns(26).Visible = False

        Frm_Estructura_Org.DGV3.Columns(27).HeaderText = "D_T_ES"
        Frm_Estructura_Org.DGV3.Columns(27).Width = 0
        Frm_Estructura_Org.DGV3.Columns(27).Visible = False

        Frm_Estructura_Org.DGV3.Columns(28).HeaderText = "No. Orden [MIXTO]"
        Frm_Estructura_Org.DGV3.Columns(28).Width = 50
        Frm_Estructura_Org.DGV3.Columns(28).Visible = True
        Frm_Estructura_Org.DGV3.Columns(28).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Estructura_Org.DGV3.Columns(28).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Estructura_Org.DGV3.Columns(28).DefaultCellStyle.BackColor = Color.Khaki

        Frm_Estructura_Org.DGV3.Columns(29).HeaderText = "No. Orden [MILITAR]"
        Frm_Estructura_Org.DGV3.Columns(29).Width = 50
        Frm_Estructura_Org.DGV3.Columns(29).Visible = True
        Frm_Estructura_Org.DGV3.Columns(29).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Estructura_Org.DGV3.Columns(29).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Estructura_Org.DGV3.Columns(29).DefaultCellStyle.BackColor = Color.LightGreen

        Frm_Estructura_Org.DGV3.Columns(30).HeaderText = "No. Orden [PAME]"
        Frm_Estructura_Org.DGV3.Columns(30).Width = 50
        Frm_Estructura_Org.DGV3.Columns(30).Visible = True
        Frm_Estructura_Org.DGV3.Columns(30).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Estructura_Org.DGV3.Columns(30).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Estructura_Org.DGV3.Columns(30).DefaultCellStyle.BackColor = Color.LightBlue

        Frm_Estructura_Org.DGV3.Columns(31).HeaderText = "ID_CARGO_ES"
        Frm_Estructura_Org.DGV3.Columns(31).Width = 0
        Frm_Estructura_Org.DGV3.Columns(31).Visible = False

        Frm_Estructura_Org.DGV3.Columns(32).HeaderText = "Cargo"
        Frm_Estructura_Org.DGV3.Columns(32).Width = 180
        Frm_Estructura_Org.DGV3.Columns(32).Visible = True
        Frm_Estructura_Org.DGV3.Columns(32).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(33).HeaderText = "ID_GP"
        Frm_Estructura_Org.DGV3.Columns(33).Width = 0
        Frm_Estructura_Org.DGV3.Columns(33).Visible = False

        Frm_Estructura_Org.DGV3.Columns(34).HeaderText = "Grado Plantilla"
        Frm_Estructura_Org.DGV3.Columns(34).Width = 80
        Frm_Estructura_Org.DGV3.Columns(34).Visible = True
        Frm_Estructura_Org.DGV3.Columns(34).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(35).HeaderText = "ID_GR"
        Frm_Estructura_Org.DGV3.Columns(35).Width = 0
        Frm_Estructura_Org.DGV3.Columns(35).Visible = False

        Frm_Estructura_Org.DGV3.Columns(36).HeaderText = "Grado Real"
        Frm_Estructura_Org.DGV3.Columns(36).Width = 80
        Frm_Estructura_Org.DGV3.Columns(36).Visible = True
        Frm_Estructura_Org.DGV3.Columns(36).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(37).HeaderText = "ID_CAT_SALARIAL"
        Frm_Estructura_Org.DGV3.Columns(37).Width = 0
        Frm_Estructura_Org.DGV3.Columns(37).Visible = False

        Frm_Estructura_Org.DGV3.Columns(38).HeaderText = "D_CAT_SALARIAL"
        Frm_Estructura_Org.DGV3.Columns(38).Width = 0
        Frm_Estructura_Org.DGV3.Columns(38).Visible = False

        Frm_Estructura_Org.DGV3.Columns(39).HeaderText = "ID_M_P"
        Frm_Estructura_Org.DGV3.Columns(39).Width = 0
        Frm_Estructura_Org.DGV3.Columns(39).Visible = False

        Frm_Estructura_Org.DGV3.Columns(40).HeaderText = "Codigo"
        Frm_Estructura_Org.DGV3.Columns(40).Width = 80
        Frm_Estructura_Org.DGV3.Columns(40).Visible = True
        Frm_Estructura_Org.DGV3.Columns(40).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(41).HeaderText = "Apellidos y Nombres"
        Frm_Estructura_Org.DGV3.Columns(41).Width = 220
        Frm_Estructura_Org.DGV3.Columns(41).Visible = True
        Frm_Estructura_Org.DGV3.Columns(41).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Frm_Estructura_Org.DGV3.Columns(41).DefaultCellStyle.BackColor = Color.Khaki

        Frm_Estructura_Org.DGV3.Columns(42).HeaderText = "Nombres y Apellidos"
        Frm_Estructura_Org.DGV3.Columns(42).Width = 0
        Frm_Estructura_Org.DGV3.Columns(42).Visible = False

        Frm_Estructura_Org.DGV3.Columns(43).HeaderText = "ID_SITUACION"
        Frm_Estructura_Org.DGV3.Columns(43).Width = 0
        Frm_Estructura_Org.DGV3.Columns(43).Visible = False

        Frm_Estructura_Org.DGV3.Columns(44).HeaderText = "Situacion"
        Frm_Estructura_Org.DGV3.Columns(44).Width = 60
        Frm_Estructura_Org.DGV3.Columns(44).Visible = True
        Frm_Estructura_Org.DGV3.Columns(44).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(45).HeaderText = "Jefe"
        Frm_Estructura_Org.DGV3.Columns(45).Width = 40
        Frm_Estructura_Org.DGV3.Columns(45).Visible = True
        Frm_Estructura_Org.DGV3.Columns(45).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(46).HeaderText = "Mixta"
        Frm_Estructura_Org.DGV3.Columns(46).Width = 0
        Frm_Estructura_Org.DGV3.Columns(46).Visible = False
        Frm_Estructura_Org.DGV3.Columns(46).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(47).HeaderText = "Militar"
        Frm_Estructura_Org.DGV3.Columns(47).Width = 40
        Frm_Estructura_Org.DGV3.Columns(47).Visible = True
        Frm_Estructura_Org.DGV3.Columns(47).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(48).HeaderText = "Pame"
        Frm_Estructura_Org.DGV3.Columns(48).Width = 40
        Frm_Estructura_Org.DGV3.Columns(48).Visible = True
        Frm_Estructura_Org.DGV3.Columns(48).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(49).HeaderText = "ID_ESTABLECIMIENTO"
        Frm_Estructura_Org.DGV3.Columns(49).Width = 0
        Frm_Estructura_Org.DGV3.Columns(49).Visible = False

        Frm_Estructura_Org.DGV3.Columns(50).HeaderText = "Establecimiento"
        Frm_Estructura_Org.DGV3.Columns(50).Width = 120
        Frm_Estructura_Org.DGV3.Columns(50).Visible = True
        Frm_Estructura_Org.DGV3.Columns(50).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Frm_Estructura_Org.DGV3.Columns(51).HeaderText = "D_ADMIN"
        Frm_Estructura_Org.DGV3.Columns(51).Width = 0
        Frm_Estructura_Org.DGV3.Columns(51).Visible = False
    End Sub
End Class