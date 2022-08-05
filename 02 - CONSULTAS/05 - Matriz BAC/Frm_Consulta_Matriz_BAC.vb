Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Consulta_Matriz_BAC
    Dim FECHA1, FECHA2 As String
    Dim F1, F2 As String
    Private Sub Frm_Consulta_Matriz_BAC_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.F5) Then
            If Me.CheckBox1.Checked = True Then
                FECHA1 = Mid(Me.DateTimePicker1.Value, 1, 10)
                F1 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA1 + "', 105), 23)"
                FECHA2 = Mid(Me.DateTimePicker2.Value, 1, 10)
                F2 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA2 + "', 105), 23)"
                CADENAsql = "SELECT * FROM [VISTA MAESTRO BAC] WHERE FECHA_INGRESO BETWEEN " & F1 & " AND " & F2 & " ORDER BY ORDEN_ESTRUCTURA"
            Else
                CADENAsql = "SELECT * FROM [VISTA MAESTRO BAC] ORDER BY ORDEN_ESTRUCTURA"
            End If
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
        End If
        If (e.KeyCode = Keys.Escape) Then
            CONSULTA_MATRIZ_BAC = False
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Consulta_Matriz_BAC_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CONSULTA_MATRIZ_BAC = True
        If Me.CheckBox1.Checked = True Then
            FECHA1 = Mid(Me.DateTimePicker1.Value, 1, 10)
            F1 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA1 + "', 105), 23)"
            FECHA2 = Mid(Me.DateTimePicker2.Value, 1, 10)
            F2 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA2 + "', 105), 23)"
            CADENAsql = "SELECT * FROM [VISTA MAESTRO BAC] WHERE FECHA_INGRESO BETWEEN " & F1 & " AND " & F2 & " ORDER BY ORDEN_ESTRUCTURA"
        Else
            CADENAsql = "SELECT * FROM [VISTA MAESTRO BAC] ORDER BY ORDEN_ESTRUCTURA"
        End If
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        CONSULTA_MATRIZ_BAC = False
        Me.Close()
    End Sub
    Private Sub CARGAR_DATAGRIDVIEW()
        Me.DGV.DataSource = Nothing
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO BAC]")
        DGV.DataSource = MiDataSet.Tables("[VISTA MAESTRO BAC]")
        'Label1.Text = Format(CType(DGV.RowCount, Decimal), "#,##0")
        Me.Label1.Text = DGV.RowCount & " Registro(s)"
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW()
        Me.DGV.Columns(0).HeaderText = "Ord."
        Me.DGV.Columns(0).Width = 30
        Me.DGV.Columns(0).Visible = True
        Me.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        'Me.DGV.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(1).HeaderText = "Actualizado"
        Me.DGV.Columns(1).Width = 0
        Me.DGV.Columns(1).Visible = False

        Me.DGV.Columns(2).HeaderText = "Primer Nombre"
        Me.DGV.Columns(2).Width = 80
        Me.DGV.Columns(2).Visible = True
        Me.DGV.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV.Columns(2).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(3).HeaderText = "Segundo Nombre"
        Me.DGV.Columns(3).Width = 80
        Me.DGV.Columns(3).Visible = True
        Me.DGV.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV.Columns(3).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(4).HeaderText = "Primer Apellido"
        Me.DGV.Columns(4).Width = 80
        Me.DGV.Columns(4).Visible = True
        Me.DGV.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV.Columns(4).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(5).HeaderText = "Segundo Apellido"
        Me.DGV.Columns(5).Width = 80
        Me.DGV.Columns(5).Visible = True
        Me.DGV.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV.Columns(5).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(6).HeaderText = "Sexo"
        Me.DGV.Columns(6).Width = 50
        Me.DGV.Columns(6).Visible = True
        Me.DGV.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV.Columns(5).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(7).HeaderText = "Estado Civil"
        Me.DGV.Columns(7).Width = 80
        Me.DGV.Columns(7).Visible = True
        Me.DGV.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV.Columns(7).DefaultCellStyle.BackColor = Color.Black
        Me.DGV.Columns(7).Frozen = True

        Me.DGV.Columns(8).HeaderText = "Conyuge"
        Me.DGV.Columns(8).Width = 200
        Me.DGV.Columns(8).Visible = True
        Me.DGV.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV.Columns(8).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(9).HeaderText = "Celular Conyuge"
        Me.DGV.Columns(9).Width = 80
        Me.DGV.Columns(9).Visible = True
        Me.DGV.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV.Columns(9).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(10).HeaderText = "Cantidad Dependientes"
        Me.DGV.Columns(10).Width = 60
        Me.DGV.Columns(10).Visible = True
        Me.DGV.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        'Me.DGV.Columns(10).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(11).HeaderText = "Ocupacion"
        Me.DGV.Columns(11).Width = 160
        Me.DGV.Columns(11).Visible = True
        Me.DGV.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV.Columns(11).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(12).HeaderText = "Profesion/ Oficio"
        Me.DGV.Columns(12).Width = 160
        Me.DGV.Columns(12).Visible = True
        Me.DGV.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV.Columns(12).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(13).HeaderText = "Fecha Ingreso"
        Me.DGV.Columns(13).Width = 80
        Me.DGV.Columns(13).Visible = False
        Me.DGV.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(14).HeaderText = "Antiguedad"
        Me.DGV.Columns(14).Width = 60
        Me.DGV.Columns(14).Visible = True
        Me.DGV.Columns(14).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(14).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV.Columns(14).DefaultCellStyle.Format = "##,##0.00"
        'Me.DGV.Columns(14).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(15).HeaderText = "Ingreso"
        Me.DGV.Columns(15).Width = 100
        Me.DGV.Columns(15).Visible = True
        Me.DGV.Columns(15).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(15).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV.Columns(15).DefaultCellStyle.Format = "##,##0.00"
        'Me.DGV.Columns(15).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(16).HeaderText = "Referencia"
        Me.DGV.Columns(16).Width = 70
        Me.DGV.Columns(16).Visible = True
        Me.DGV.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(16).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV.Columns(16).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(17).HeaderText = "Beneficiario"
        Me.DGV.Columns(17).Width = 100
        Me.DGV.Columns(17).Visible = True
        Me.DGV.Columns(17).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(17).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV.Columns(17).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(18).HeaderText = "Fecha Nacimiento"
        Me.DGV.Columns(18).Width = 80
        Me.DGV.Columns(18).Visible = True
        Me.DGV.Columns(18).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(18).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        'Me.DGV.Columns(18).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(19).HeaderText = "Pais Nacimiento"
        Me.DGV.Columns(19).Width = 80
        Me.DGV.Columns(19).Visible = True
        Me.DGV.Columns(19).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(19).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV.Columns(19).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(20).HeaderText = "Nacionalidad"
        Me.DGV.Columns(20).Width = 80
        Me.DGV.Columns(20).Visible = True
        Me.DGV.Columns(20).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(20).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV.Columns(19).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(21).HeaderText = "Pais Residencia"
        Me.DGV.Columns(21).Width = 80
        Me.DGV.Columns(21).Visible = True
        Me.DGV.Columns(21).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(21).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV.Columns(21).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(22).HeaderText = "Direccion Domiciliar"
        Me.DGV.Columns(22).Width = 250
        Me.DGV.Columns(22).Visible = True
        Me.DGV.Columns(22).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(22).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV.Columns(22).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(23).HeaderText = "Municipio Domiciliar"
        Me.DGV.Columns(23).Width = 80
        Me.DGV.Columns(23).Visible = True
        Me.DGV.Columns(23).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(23).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV.Columns(23).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(24).HeaderText = "Departamento Domiciliar"
        Me.DGV.Columns(24).Width = 80
        Me.DGV.Columns(24).Visible = True
        Me.DGV.Columns(24).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(24).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV.Columns(24).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(25).HeaderText = "Telefono"
        Me.DGV.Columns(25).Width = 80
        Me.DGV.Columns(25).Visible = True
        Me.DGV.Columns(25).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(25).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV.Columns(25).DefaultCellStyle.BackColor = Color.Black


        Me.DGV.Columns(26).HeaderText = "Celular"
        Me.DGV.Columns(26).Width = 80
        Me.DGV.Columns(26).Visible = True
        Me.DGV.Columns(26).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(26).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV.Columns(26).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(27).HeaderText = "Correo Electronico"
        Me.DGV.Columns(27).Width = 150
        Me.DGV.Columns(27).Visible = True
        Me.DGV.Columns(27).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(27).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV.Columns(27).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(28).HeaderText = "Tipo Identificacion"
        Me.DGV.Columns(28).Width = 100
        Me.DGV.Columns(28).Visible = True
        Me.DGV.Columns(28).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(28).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV.Columns(29).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(29).HeaderText = "Id"
        Me.DGV.Columns(29).Width = 100
        Me.DGV.Columns(29).Visible = True
        Me.DGV.Columns(29).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(29).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        'Me.DGV.Columns(29).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(30).HeaderText = "Pais Emision Id"
        Me.DGV.Columns(30).Width = 100
        Me.DGV.Columns(30).Visible = True
        Me.DGV.Columns(30).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(30).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        'Me.DGV.Columns(30).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(31).HeaderText = "Fecha Emision Id"
        Me.DGV.Columns(31).Width = 80
        Me.DGV.Columns(31).Visible = True
        Me.DGV.Columns(31).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(31).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        'Me.DGV.Columns(31).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(32).HeaderText = "Fecha Vence Id"
        Me.DGV.Columns(32).Width = 80
        Me.DGV.Columns(32).Visible = True
        Me.DGV.Columns(32).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(32).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        'Me.DGV.Columns(32).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(33).HeaderText = "Nivel 1"
        Me.DGV.Columns(33).Width = 100
        Me.DGV.Columns(33).Visible = True
        Me.DGV.Columns(33).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(33).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV.Columns(33).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(34).HeaderText = "Nivel 2"
        Me.DGV.Columns(34).Width = 100
        Me.DGV.Columns(34).Visible = True
        Me.DGV.Columns(34).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(34).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV.Columns(34).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(35).HeaderText = "Nivel 3"
        Me.DGV.Columns(35).Width = 100
        Me.DGV.Columns(35).Visible = True
        Me.DGV.Columns(35).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(35).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV.Columns(35).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(36).HeaderText = "Nivel 4"
        Me.DGV.Columns(36).Width = 100
        Me.DGV.Columns(36).Visible = True
        Me.DGV.Columns(36).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(36).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV.Columns(36).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(37).HeaderText = "Nivel 5"
        Me.DGV.Columns(37).Width = 100
        Me.DGV.Columns(37).Visible = True
        Me.DGV.Columns(37).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(37).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV.Columns(37).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(38).HeaderText = "Nivel 6"
        Me.DGV.Columns(38).Width = 100
        Me.DGV.Columns(38).Visible = True
        Me.DGV.Columns(38).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(38).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV.Columns(38).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(39).HeaderText = "Nivel 7"
        Me.DGV.Columns(39).Width = 100
        Me.DGV.Columns(39).Visible = True
        Me.DGV.Columns(39).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(39).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV.Columns(39).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(40).HeaderText = "Nivel 8"
        Me.DGV.Columns(40).Width = 100
        Me.DGV.Columns(40).Visible = True
        Me.DGV.Columns(40).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(40).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV.Columns(40).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(41).HeaderText = "Grado"
        Me.DGV.Columns(41).Width = 100
        Me.DGV.Columns(41).Visible = True
        Me.DGV.Columns(41).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(41).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DGV.Columns(41).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(42).HeaderText = "Militar"
        Me.DGV.Columns(42).Width = 50
        Me.DGV.Columns(42).Visible = True
        Me.DGV.Columns(42).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(42).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        'Me.DGV.Columns(42).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(43).HeaderText = "PAME"
        Me.DGV.Columns(43).Width = 50
        Me.DGV.Columns(43).Visible = True
        Me.DGV.Columns(43).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(43).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        'Me.DGV.Columns(43).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(44).HeaderText = "Jefe"
        Me.DGV.Columns(44).Width = 50
        Me.DGV.Columns(44).Visible = True
        Me.DGV.Columns(44).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(44).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        'Me.DGV.Columns(44).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(45).HeaderText = "Orden Estructura"
        Me.DGV.Columns(45).Width = 0
        Me.DGV.Columns(45).Visible = False
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call VERIFICAR_ACCESOS("305") : If HAY_ACCESO = False Then : Exit Sub : End If
        MENSAJE = MsgBox("Se dispone a Actualizar la Lista de Empleados, los registros marcados como <ACTUALIZADO> no seran afectados, ¿Esta seguro de Continuar?", vbInformation + vbYesNo, "Mensaje del Sistema")
        If MENSAJE = vbYes Then
            Me.Button1.Enabled = False
            Me.Button2.Enabled = False
            Me.Button5.Enabled = False
            Call BORRAR_MAESTRO_BAC()
            Call INICIAR_MAESTRO_BAC()
            Frm_pBACLIST_Actualizar.ShowDialog()
            If Me.CheckBox1.Checked = True Then
                FECHA1 = Mid(Me.DateTimePicker1.Value, 1, 10)
                F1 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA1 + "', 105), 23)"
                FECHA2 = Mid(Me.DateTimePicker2.Value, 1, 10)
                F2 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA2 + "', 105), 23)"
                CADENAsql = "SELECT * FROM [VISTA MAESTRO BAC] WHERE FECHA_INGRESO BETWEEN " & F1 & " AND " & F2 & " ORDER BY ORDEN_ESTRUCTURA"
            Else
                CADENAsql = "SELECT * FROM [VISTA MAESTRO BAC] ORDER BY ORDEN_ESTRUCTURA"
            End If
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Me.Button1.Enabled = True
            Me.Button2.Enabled = True
            Me.Button5.Enabled = True
        End If
    End Sub
    Private Sub INICIAR_MAESTRO_BAC()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "DBCC CHECKIDENT ([MAESTRO BAC], RESEED, 0)"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub BORRAR_MAESTRO_BAC()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiSqlCommand As SqlCommand = Conexion.CxRRHH.CreateCommand
        Try
            MiSqlCommand.CommandText = "DELETE FROM [MAESTRO BAC]"
            MiSqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Call VERIFICAR_ACCESOS("306") : If HAY_ACCESO = False Then : Exit Sub : End If
        Try
            TITULO_EXCEL = "LISTA DE EMPLEADOS"
            EXPORTAR_DATOS_A_EXCEL(DGV, "FECHA: " & Mid(Now, 1, 10))
        Catch ex As Exception
            MessageBox.Show("No se puede generar el documento Excel.", "Mensaje del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Me.DateTimePicker1.Enabled = Me.CheckBox1.Checked
        Me.DateTimePicker2.Enabled = Me.CheckBox1.Checked
        If Me.CheckBox1.Checked = True Then
            FECHA1 = Mid(Me.DateTimePicker1.Value, 1, 10)
            F1 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA1 + "', 105), 23)"
            FECHA2 = Mid(Me.DateTimePicker2.Value, 1, 10)
            F2 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA2 + "', 105), 23)"
            CADENAsql = "SELECT * FROM [VISTA MAESTRO BAC] WHERE FECHA_INGRESO BETWEEN " & F1 & " AND " & F2 & " ORDER BY ORDEN_ESTRUCTURA"
        Else
            CADENAsql = "SELECT * FROM [VISTA MAESTRO BAC] ORDER BY ORDEN_ESTRUCTURA"
        End If
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
    End Sub
    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        If Me.CheckBox1.Checked = True Then
            FECHA1 = Mid(Me.DateTimePicker1.Value, 1, 10)
            F1 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA1 + "', 105), 23)"
            FECHA2 = Mid(Me.DateTimePicker2.Value, 1, 10)
            F2 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA2 + "', 105), 23)"
            CADENAsql = "SELECT * FROM [VISTA MAESTRO BAC] WHERE FECHA_INGRESO BETWEEN " & F1 & " AND " & F2 & " ORDER BY ORDEN_ESTRUCTURA"
        Else
            CADENAsql = "SELECT * FROM [VISTA MAESTRO BAC] ORDER BY ORDEN_ESTRUCTURA"
        End If
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
    End Sub
    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        If Me.CheckBox1.Checked = True Then
            FECHA1 = Mid(Me.DateTimePicker1.Value, 1, 10)
            F1 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA1 + "', 105), 23)"
            FECHA2 = Mid(Me.DateTimePicker2.Value, 1, 10)
            F2 = "Convert(VARCHAR(10), Convert(Date, '" + FECHA2 + "', 105), 23)"
            CADENAsql = "SELECT * FROM [VISTA MAESTRO BAC] WHERE FECHA_INGRESO BETWEEN " & F1 & " AND " & F2 & " ORDER BY ORDEN_ESTRUCTURA"
        Else
            CADENAsql = "SELECT * FROM [VISTA MAESTRO BAC] ORDER BY ORDEN_ESTRUCTURA"
        End If
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
    End Sub
End Class